"""FastAPI backend for the PyPaxTagTool webapp.

Implements the frozen REST contract (see the Phase 3 brief): connect,
PLC status, AOI list/read, TSV paste, staged changes, guarded writes,
audit log, and spreadsheet export. All state lives in one
:class:`AppState` guarded by a lock; one PLC connection at a time.

The backend reads through the same ``pypax`` core library as the CLI:
instances are discovered from the PLC, columns come from the template
workbook, reads use :func:`pypax.plc.read_plc_row`, and writes use the
ordered :func:`pypax.diff.row_differences` plus
:func:`pypax.plc.write_plc_row`. Errors are JSON ``{"detail": ...}``;
tracebacks never leak.
"""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pycomm3 import LogixDriver

from ..aoi import load_aoi_definitions
from ..config import DEFAULT_AOI_VERSION, AppSettings, load_app_config
from ..custom_aoi import (
    DEFAULT_CUSTOM_DIR,
    check_collision,
    confirmed_meta_path,
    diff_columns,
    hash_file,
    load_confirmed_meta,
    load_custom_aois,
    read_source_sha256,
    save_custom_aoi,
)
from ..diff import READBACK_ABSOLUTE_TOLERANCE, row_differences, values_equal
from ..export import build_export_workbook
from ..plc import MAX_TAGS_PER_READ, make_tag_list, read_plc_row, write_plc_row
from ..snapshots import (
    AUTO_KEEP_PER_PLC,
    build_snapshot,
    diff_snapshot,
    find_snapshot_file,
    list_snapshots,
    load_snapshot,
    parse_snapshot_bytes,
    prune_auto_snapshots,
    resolve_snapshot_dir,
    sanitize_snapshot_name,
    save_snapshot,
    snapshot_counts,
)
from ..string_sizes import merge_learned_sizes, merged_sizes
from .aoi_model import discover_instances, load_combined_model, read_aoi_rows
from .import_aoi import (
    build_proposal,
    check_upload_target,
    consume_proposal,
    parse_l5x_bytes,
    parse_multipart_upload,
    register_proposal,
    resolve_import_dir,
    save_upload,
    scan_import_dir,
    valid_aoi_name,
)
from .paste import coerce_value, parse_tsv
from .state import AppState

logger = logging.getLogger(__name__)

DEFAULT_TEMPLATE_NAME = "ProcessLibraryOnlineConfigTool.xlsm"

_PLACEHOLDER_PAGE = """<!doctype html>
<html><head><meta charset="utf-8"><title>PyPaxTagTool</title></head>
<body><h1>PyPaxTagTool</h1>
<p>The webapp backend is running. The frontend is not built yet.</p>
<p>API: <a href="/docs">/docs</a> &middot; <a href="/api/health">/api/health</a></p>
</body></html>
"""


class ConnectRequest(BaseModel):
    commpath: str
    profile: str | None = None


class PasteRequest(BaseModel):
    tsv: str
    # Where the pasted block starts: top-left cell maps to
    # (anchor_instance, anchor_subtag); cells flow right across subtags
    # and down across instances. Absent: first instance, first subtag.
    anchor_instance: str | None = None
    anchor_subtag: str | None = None


class WriteRequest(BaseModel):
    plc_name: str


class ConfirmRequest(BaseModel):
    proposal_id: str
    # Subtags to keep, in any order; None keeps every proposed column.
    selected: list[str] | None = None


class SnapshotRequest(BaseModel):
    # Manual snapshot name ([A-Za-z0-9_-]+); None means an auto-style
    # <plc>_<YYYYMMDD-HHMMSS> name.
    name: str | None = None


def default_plc_factory(commpath):
    """Build a real pycomm3 driver for a comm path."""
    return LogixDriver(commpath, init_tags=True, init_program_tags=True)


def _export_filename(plc_name):
    """``<plc_name>_ConfigTags.xlsx`` — the CLI write mode requires the
    PLC name in the file name. Anything outside ``[A-Za-z0-9._-]`` is
    replaced so the name is safe in a download header; ``None`` (not
    connected) falls back to ``PyPax_ConfigTags.xlsx``."""
    base = plc_name or "PyPax"
    safe = "".join(c if c.isalnum() or c in "._-" else "_" for c in base)
    return f"{safe}_ConfigTags.xlsx"


def _resolve_template_path(settings: AppSettings, template_path):
    """Explicit path, else the setting, else the template in the cwd."""
    if template_path is not None:
        return str(template_path)
    if settings.template_path:
        return settings.template_path
    candidate = Path.cwd() / DEFAULT_TEMPLATE_NAME
    if candidate.exists():
        return str(candidate)
    return None


def _require_connected(state: AppState) -> None:
    if not state.connected:
        raise HTTPException(status_code=409, detail="not connected to a PLC")


def _ensure_model(state: AppState) -> dict:
    """Template ∪ custom AOI model, cached per connection; 503 when unavailable."""
    if state.aoi_model is None:
        if state.template_path is None:
            raise HTTPException(
                status_code=503,
                detail="template workbook not found: set template_path or place "
                f"{DEFAULT_TEMPLATE_NAME} in the working directory",
            )
        try:
            model, custom_names = load_combined_model(
                state.template_path, state.custom_aoi_dir, state.column_meta
            )
        except FileNotFoundError:
            raise HTTPException(
                status_code=503,
                detail=f"template workbook not found: {state.template_path}",
            )
        except Exception as exc:  # unreadable workbook / corrupt custom store
            raise HTTPException(status_code=503, detail=f"AOI model could not be loaded: {exc}")
        state.aoi_model = model
        state.custom_aoi_names = custom_names
    return state.aoi_model


def _load_custom_names(state: AppState) -> set[str]:
    """Custom AOI names, without needing a PLC connection or a template."""
    try:
        return set(load_custom_aois(state.custom_aoi_dir))
    except ValueError as exc:
        raise HTTPException(status_code=503, detail=f"custom AOI store could not be read: {exc}")


def _canonical_custom_name(state: AppState, name: str) -> str | None:
    """The stored custom-AOI name matching ``name`` case-insensitively.

    ControlLogix tag types are case-insensitive, and on a
    case-insensitive filesystem ``P_Foo.yaml`` and ``p_foo.yaml`` are
    the same file — a re-import spelled in a different case is the
    same AOI, not a new one. Returns ``None`` when ``name`` is not an
    imported custom AOI.
    """
    lowered = name.lower()
    for custom in _load_custom_names(state):
        if custom.lower() == lowered:
            return custom
    return None


def _before_columns(state: AppState, canonical: str) -> dict:
    """Confirmed columns for a custom AOI, for diffing a re-import.

    Subtags come from the stored YAML (the confirmed definition);
    types/descriptions/defaults come from the confirmed-metadata
    cache, which is written only at confirm time — never by a
    proposal registration, so an unconfirmed re-proposal cannot move
    the "before" side. When the cache misses (a restart, or the
    source file is gone), the sidecar written next to the YAML at
    confirm time is the durable fallback. A subtag with no metadata
    anywhere maps to ``None`` — "known to exist, details unknown" —
    and :func:`pypax.custom_aoi.diff_columns` counts it unchanged
    rather than guessing. (The cache is keyed by confirmed spelling,
    so the lookup is case-insensitive too.)
    """
    try:
        subtags = load_custom_aois(state.custom_aoi_dir).get(canonical, [])
    except ValueError:
        subtags = []
    meta = {}
    for key, value in state.confirmed_meta.items():
        if key.lower() == canonical.lower():
            meta = value
            break
    if not meta:
        yaml_path = _find_custom_aoi_file(state, canonical)
        if yaml_path is not None:
            meta = load_confirmed_meta(yaml_path)
    return {subtag: meta.get(subtag) for subtag in subtags}


def _reimport_diff(state: AppState, canonical: str, new_columns) -> dict:
    """Diff a new proposal's columns against the confirmed AOI.

    ``before_details_known`` is False when the confirmed column
    details are not available anywhere (a legacy import with no
    sidecar): type/description/default changes cannot be detected
    then, and the diff must say so instead of reporting those columns
    unchanged.
    """
    before = _before_columns(state, canonical)
    diff = diff_columns(before, list(new_columns))
    diff["before_details_known"] = all(column is not None for column in before.values())
    return diff


def _source_hash(proposal) -> str | None:
    """SHA-256 of a proposal's source file, or None when it is gone."""
    if proposal.source_path is None:
        return None
    try:
        return hash_file(proposal.source_path)
    except OSError:
        return None


def _confirm_update(state: AppState, proposal, canonical: str) -> dict:
    """Confirm an update proposal: rewrite the custom AOI in place.

    All-or-nothing — the client's ``selected`` list is ignored; the
    diff the proposal carried is what the user reviewed. Refuses with
    409 when staged changes exist on the AOI: a re-import must not
    silently drop staged edits — discard or write them first.
    """
    if any(key[0].lower() == canonical.lower() for key in state.staged):
        raise HTTPException(
            status_code=409,
            detail=(
                f"AOI {canonical!r} has staged changes; discard or write them before re-importing"
            ),
        )
    diff = _reimport_diff(state, canonical, proposal.columns)
    target = _find_custom_aoi_file(state, canonical)
    try:
        save_custom_aoi(
            canonical,
            proposal.columns,
            proposal.source_file,
            proposal.revision,
            path=target if target is not None else state.custom_aoi_dir,
            source_sha256=_source_hash(proposal),
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    # Column metadata is keyed by the confirmed (canonical) name — the
    # proposal may spell the AOI in a different case. The confirmed
    # cache is what later re-import diffs compare against.
    state.column_meta[canonical] = {column.subtag: column for column in proposal.columns}
    state.confirmed_meta[canonical] = {column.subtag: column for column in proposal.columns}
    if proposal.aoi != canonical:
        state.column_meta.pop(proposal.aoi, None)
    # Learn string-UDT sizes from the import (pinned sizes win;
    # conflicts come back as warnings, never as failures) — the same
    # merge a fresh import does. A re-import must not silently drop
    # what its file taught.
    warnings = merge_learned_sizes(
        proposal.learned_sizes,
        proposal.source_file,
        path=state.string_registry_path,
    )
    consume_proposal(state, proposal.proposal_id)
    state.aoi_model = None
    state.drop_aoi_data(canonical)
    state.audit.record(
        {
            "timestamp": datetime.now(timezone.utc).isoformat(),
            "plc_name": state.plc_name,
            "action": "update_custom_aoi",
            "tags": len(proposal.columns),
            "results": {
                "aoi": canonical,
                "columns": len(proposal.columns),
                "added": len(diff["added"]),
                "removed": len(diff["removed"]),
                "changed": len(diff["changed"]),
            },
        }
    )
    response = {"aoi": canonical, "columns": len(proposal.columns), "updated": True}
    if warnings:
        response["warnings"] = warnings
    return response


def _find_custom_aoi_file(state: AppState, name: str):
    """The custom YAML file defining ``name``, or None.

    One AOI per file in practice (that is what the importer writes); a
    hand-made file defining several AOIs is still found correctly.
    """
    directory = Path(state.custom_aoi_dir)
    if not directory.is_dir():
        return None
    for yaml_path in sorted(directory.glob("*.yaml")):
        try:
            definitions = load_aoi_definitions(yaml_path)
        except ValueError:
            continue
        if name in definitions:
            return yaml_path
    return None


def _maybe_seed_confirmed_meta(state: AppState, path, proposal) -> None:
    """Seed the confirmed-column cache at startup when the file on
    disk is the confirmed import — its hash matches the one stored in
    the YAML. A changed file (or a YAML without a hash) seeds nothing:
    the re-import diff then falls back to added/removed. Never raises
    and never blocks startup.
    """
    try:
        custom = load_custom_aois(state.custom_aoi_dir)
    except ValueError:
        return
    canonical = next((name for name in custom if name.lower() == proposal.aoi.lower()), None)
    if canonical is None:
        return
    yaml_path = _find_custom_aoi_file(state, canonical)
    if yaml_path is None:
        return
    stored = read_source_sha256(yaml_path)
    if stored is None:
        return
    try:
        current = hash_file(path)
    except OSError:
        return
    if stored == current:
        state.confirmed_meta[canonical] = {column.subtag: column for column in proposal.columns}


def _startup_scan(state: AppState) -> None:
    """Parse the import directory into proposals at startup.

    Never auto-confirms anything: a proposal is inert until the user
    confirms it. A broken import directory must never kill startup.
    """
    try:
        results = scan_import_dir(state.import_dir)
    except OSError as exc:
        logger.warning("AOI import scan failed for %s: %s", state.import_dir, exc)
        return
    proposed = skipped = 0
    for result in results:
        if result.error is not None:
            logger.warning("AOI import scan: skipping %s: %s", result.path.name, result.error)
            skipped += 1
            continue
        proposal = build_proposal(result.definition, result.path.name, result.path)
        register_proposal(state, proposal)
        _maybe_seed_confirmed_meta(state, result.path, proposal)
        proposed += 1
    logger.info(
        "AOI import scan: %d .L5X file(s) in %s (%d proposals, %d skipped)",
        len(results),
        state.import_dir,
        proposed,
        skipped,
    )


def _ensure_baseline(state: AppState, aoi: str) -> None:
    """Read an AOI's rows from the PLC into the baseline (once per connection)."""
    model = _ensure_model(state)
    if aoi not in model:
        raise HTTPException(status_code=404, detail=f"unknown AOI {aoi!r}")
    if aoi in state.baseline:
        return
    aoi_model = model[aoi]
    aoi_model.instances = discover_instances(state.plc, aoi)
    try:
        state.baseline[aoi] = read_aoi_rows(state.plc, aoi_model)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"failed to read {aoi} from PLC: {exc}")


def _read_back_values(plc, tags):
    """Chunked read-back of written tags: ``{tag: formatted value}``.

    Reads in chunks of at most ``MAX_TAGS_PER_READ`` so no request
    exceeds the CIP packet limit. A tag whose read fails (or whose whole
    chunk raises) maps to ``None`` — the caller treats that as an
    unverified read-back, never as a crash.
    """
    values = {}
    for start in range(0, len(tags), MAX_TAGS_PER_READ):
        chunk = tags[start : start + MAX_TAGS_PER_READ]
        try:
            formatted, _raw = read_plc_row(plc, chunk)
        except Exception:
            for tag in chunk:
                values[tag] = None
            continue
        for tag, value in formatted:
            values[tag] = value
    return values


# ----------------------------------------------------------------------
# PLC snapshots (Phase 9a)
# ----------------------------------------------------------------------


def _aoi_version(state: AppState) -> str:
    """AOI version for a snapshot: the connected profile's, else the default."""
    if state.profile is not None:
        return state.profile.aoi_version
    return DEFAULT_AOI_VERSION


def _resolve_snapshot_file(state: AppState, name: str) -> Path:
    """The file for a snapshot name (manual, else auto).

    404 when no snapshot with that name exists, 400 for an invalid
    name. Does NOT read the file: a corrupt snapshot must still be
    downloadable (for inspection) and deletable.
    """
    try:
        path = find_snapshot_file(state.snapshot_dir, name)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    if path is None:
        raise HTTPException(status_code=404, detail=f"unknown snapshot {name!r}")
    return path


def _load_named_snapshot(state: AppState, name: str) -> tuple[dict, Path]:
    """The snapshot dict + file for ``name``.

    404 when no snapshot with that name exists, 400 when the file is
    there but corrupt — a corrupt snapshot must never be compared or
    restored from.
    """
    path = _resolve_snapshot_file(state, name)
    try:
        return load_snapshot(path), path
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"snapshot {name!r} is corrupt: {exc}")


def _refresh_all_baselines(state: AppState) -> None:
    """Re-read every AOI's rows from the PLC into the baseline.

    Unlike ``_ensure_baseline`` (which keeps already-loaded rows),
    this refreshes them: a snapshot must hold the PLC's values
    *now*, not the cached values from when a tab was loaded — a value
    that drifted since then would otherwise be snapshotted stale and
    a later restore would silently revert the drift. Raises the same
    502 as ``_ensure_baseline`` when a read fails.
    """
    model = _ensure_model(state)
    for aoi in model:
        _refresh_baseline(state, aoi)


def _refresh_baseline(state: AppState, aoi: str) -> None:
    """Re-read an AOI's rows from the PLC, replacing the baseline.

    Compare/restore need the PLC's values *now*, not the cached
    baseline from when the tab was loaded — otherwise a value that
    moved since then would not show as a difference (and a restore
    would stage against a stale old value). Instances are
    re-discovered too, so a deleted instance drops out of the diff.
    Raises the same 502 as ``_ensure_baseline`` when a read fails.
    """
    model = _ensure_model(state)
    if aoi not in model:
        raise HTTPException(status_code=404, detail=f"unknown AOI {aoi!r}")
    aoi_model = model[aoi]
    aoi_model.instances = discover_instances(state.plc, aoi)
    try:
        state.baseline[aoi] = read_aoi_rows(state.plc, aoi_model)
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"failed to read {aoi} from PLC: {exc}")


def _take_auto_snapshot(state: AppState) -> dict:
    """Capture the pre-write baseline to ``snapshots/auto/``.

    Never raises: any failure (unreadable AOI, unwritable directory)
    is returned as ``{"ok": False, "error": ...}`` so the caller can
    log it and carry on with the write. Keeps the newest 10
    auto-snapshots per PLC, pruning older ones.
    """
    try:
        _refresh_all_baselines(state)
        snapshot = build_snapshot(state.plc_name, state.baseline, _aoi_version(state))
        saved = save_snapshot(state.snapshot_dir, snapshot)
        prune_auto_snapshots(state.snapshot_dir, state.plc_name, keep=AUTO_KEEP_PER_PLC)
        return {"ok": True, "name": saved.stem}
    except Exception as exc:
        logger.warning("auto-snapshot failed before write: %s", exc)
        return {"ok": False, "error": str(exc)}


def _restore_from_snapshot(state: AppState, snapshot: dict) -> tuple[int, list[dict]]:
    """Stage the values of a snapshot that differ from the live PLC.

    Iterates the snapshot's own values (a tag the snapshot does not
    have is simply not restored — it is not a failure). The baseline
    is refreshed from the PLC first, so "current" is what the PLC
    holds now. Returns ``(staged, skipped)``. A value is staged
    through the normal validation path
    (:func:`pypax.web.paste.coerce_value` against the column's data
    type); anything that cannot be staged is reported in ``skipped``
    with a reason and never staged. Nothing here writes to the PLC
    — restore only stages.
    """
    model = _ensure_model(state)
    for aoi in snapshot["aois"]:
        if aoi in model:
            _refresh_baseline(state, aoi)
    types = {
        aoi: {column.subtag: column.data_type for column in model[aoi].columns} for aoi in model
    }
    staged = 0
    skipped: list[dict] = []
    for aoi, instances in snapshot["aois"].items():
        for instance, values in instances.items():
            for subtag, snap_value in values.items():
                entry = {"aoi": aoi, "instance": instance, "subtag": subtag}
                if snap_value is None:
                    skipped.append({**entry, "reason": "no value in snapshot"})
                    continue
                if aoi not in model:
                    skipped.append({**entry, "reason": f"unknown AOI {aoi!r}"})
                    continue
                if subtag not in types[aoi]:
                    skipped.append({**entry, "reason": f"unknown subtag {subtag!r} in {aoi}"})
                    continue
                current_values = state.baseline.get(aoi, {}).get(instance)
                if current_values is None:
                    skipped.append({**entry, "reason": f"instance {instance!r} not found in PLC"})
                    continue
                value, reason = coerce_value(
                    str(snap_value), types[aoi][subtag], state.string_registry_path
                )
                if reason is not None:
                    skipped.append({**entry, "reason": reason})
                    continue
                current = current_values[subtag]
                if value == current:
                    continue  # coercion normalized it to the current value
                state.staged[(aoi, instance, subtag)] = {
                    "aoi": aoi,
                    "instance": instance,
                    "subtag": subtag,
                    "tag": instance + subtag,
                    "old_value": current,
                    "new_value": value,
                }
                staged += 1
    return staged, skipped


def create_app(
    plc_factory=None,
    template_path=None,
    settings=None,
    profiles=None,
    audit_dir=None,
    import_dir=None,
    custom_aoi_dir=None,
    snapshot_dir=None,
    string_registry_path=None,
):
    """Build the FastAPI app. Tests inject a fake PLC via ``plc_factory``."""
    settings = settings or AppSettings()
    profiles = profiles or {}
    state = AppState(
        plc_factory=plc_factory or default_plc_factory,
        template_path=_resolve_template_path(settings, template_path),
        settings=settings,
        profiles=profiles,
        audit_dir=audit_dir,
    )
    state.custom_aoi_dir = (
        Path(custom_aoi_dir) if custom_aoi_dir is not None else DEFAULT_CUSTOM_DIR
    )
    state.import_dir = resolve_import_dir(settings, import_dir)
    state.string_registry_path = (
        Path(string_registry_path) if string_registry_path is not None else None
    )
    state.snapshot_dir = resolve_snapshot_dir(settings, snapshot_dir)
    app = FastAPI(title="PyPaxTagTool")
    app.state.pypax = state  # tests and tooling inspect the live state here
    _startup_scan(state)

    @app.get("/api/health")
    def health():
        return {"status": "ok"}

    @app.get("/api/string-limits")
    def string_limits():
        """Merged string-size map: pinned sizes over learned sizes.

        The frontend fetches this at startup so a type learned from
        an imported ``.L5X`` gets its limit in the grid too. ``STRING_N``
        is not in the map — it holds N by the naming rule.
        """
        return merged_sizes(state.string_registry_path)

    @app.post("/api/connect")
    def connect(body: ConnectRequest):
        with state.lock:
            profile = None
            if body.profile is not None:
                profile = state.profiles.get(body.profile)
                if profile is None:
                    raise HTTPException(status_code=404, detail=f"unknown profile {body.profile!r}")
            if state.connected:
                try:
                    state.plc.close()
                except Exception:
                    pass
                state.reset_connection()
            plc = state.plc_factory(body.commpath)
            try:
                plc.open()
                plc_name = plc.get_plc_name()
            except Exception as exc:
                try:
                    plc.close()
                except Exception:
                    pass
                state.audit.record(
                    {
                        "timestamp": datetime.now(timezone.utc).isoformat(),
                        "plc_name": None,
                        "action": "connect",
                        "tags": 0,
                        "results": {"connected": False, "error": str(exc)},
                    }
                )
                return {
                    "connected": False,
                    "plc_name": None,
                    "commpath": body.commpath,
                    "error": str(exc),
                }
            state.plc = plc
            state.commpath = body.commpath
            state.plc_name = plc_name
            state.profile = profile
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": plc_name,
                    "action": "connect",
                    "tags": 0,
                    "results": {"connected": True},
                }
            )
            return {"connected": True, "plc_name": plc_name, "commpath": body.commpath}

    @app.post("/api/disconnect")
    def disconnect():
        with state.lock:
            was_connected = state.connected
            plc_name = state.plc_name
            if state.connected:
                try:
                    state.plc.close()
                except Exception:
                    pass
            state.reset_connection()
            if was_connected:
                state.audit.record(
                    {
                        "timestamp": datetime.now(timezone.utc).isoformat(),
                        "plc_name": plc_name,
                        "action": "disconnect",
                        "tags": 0,
                        "results": {"connected": False},
                    }
                )
            return {"connected": False}

    @app.get("/api/plc")
    def plc_status():
        return {
            "connected": state.connected,
            "plc_name": state.plc_name,
            "commpath": state.commpath,
        }

    @app.get("/api/aois")
    def list_aois():
        with state.lock:
            _require_connected(state)
            model = _ensure_model(state)
            result = []
            for name in model:
                instances = discover_instances(state.plc, name)
                result.append(
                    {
                        "name": name,
                        "instance_count": len(instances),
                        "subtag_count": len(model[name].columns),
                        "custom": name in state.custom_aoi_names,
                    }
                )
            return result

    @app.post("/api/aois/import")
    async def import_aoi(request: Request):
        """Upload one ``.L5X`` export and get a proposal back (no tab yet).

        The body is parsed with the stdlib ``email`` parser (no
        python-multipart dependency). The uploaded file is persisted
        into the import directory so a later scan/startup finds it
        again. 400 on a missing/malformed upload, a non-``.L5X`` name,
        an oversized file, or an unparseable export. When the AOI is
        already imported (case-insensitive), the upload is a re-import:
        it returns ``{status: "update_available", proposal_id, diff}``
        (plus the full proposal) instead of a fresh proposal — confirm
        it to rewrite the AOI in place.
        """
        body = await request.body()
        content_type = request.headers.get("content-type", "")
        with state.lock:
            try:
                filename, data = parse_multipart_upload(body, content_type)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            if Path(filename).suffix.lower() != ".l5x":
                raise HTTPException(
                    status_code=400,
                    detail=f"only .L5X files can be imported, got {filename!r}",
                )
            try:
                definition = parse_l5x_bytes(data, filename or "upload.L5X")
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            try:
                check_upload_target(state.import_dir, filename, definition.name)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            try:
                saved = save_upload(state.import_dir, filename, data)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            canonical = _canonical_custom_name(state, definition.name)
            proposal = build_proposal(definition, saved.name, saved)
            if canonical is not None:
                # Re-import: diff while column_meta still holds the
                # confirmed columns — registering the proposal
                # overwrites it with the new file's columns.
                diff = _reimport_diff(state, canonical, proposal.columns)
                register_proposal(state, proposal)
                return {
                    **proposal.to_json(),
                    "status": "update_available",
                    "update_available": True,
                    "diff": diff,
                }
            register_proposal(state, proposal)
            return proposal.to_json()

    @app.get("/api/aois/scan")
    def scan_aois():
        """Scan the import directory for ``.L5X`` exports.

        Each parseable file whose AOI is not already imported becomes a
        proposal. Files already imported are skipped with reason
        ``already imported`` — unless the file on disk changed since
        the import: those are reported under ``updates`` as full
        proposals with ``status: "update_available"`` and the column
        diff. Change is detected by the SHA-256 stored in the custom
        YAML at confirm time; a YAML without one (a legacy import)
        falls back to the column diff. Unparseable files and template
        collisions are skipped with a descriptive reason. No
        connection needed.
        """
        with state.lock:
            custom_names = _load_custom_names(state)
            proposals = []
            skipped = []
            updates = []
            for result in scan_import_dir(state.import_dir):
                if result.error is not None:
                    skipped.append(
                        {"file": result.path.name, "reason": f"unparseable: {result.error}"}
                    )
                    continue
                name = result.definition.name
                canonical = next(
                    (custom for custom in custom_names if custom.lower() == name.lower()),
                    None,
                )
                if canonical is not None:
                    yaml_path = _find_custom_aoi_file(state, canonical)
                    stored_hash = read_source_sha256(yaml_path) if yaml_path is not None else None
                    try:
                        current_hash = hash_file(result.path)
                    except OSError:
                        current_hash = None
                    proposal = build_proposal(result.definition, result.path.name, result.path)
                    diff = _reimport_diff(state, canonical, proposal.columns)
                    changed = (
                        stored_hash is not None
                        and current_hash is not None
                        and stored_hash != current_hash
                    ) or (
                        stored_hash is None
                        and (diff["added"] or diff["removed"] or diff["changed"])
                    )
                    if not changed:
                        skipped.append({"file": result.path.name, "reason": "already imported"})
                        continue
                    register_proposal(state, proposal)
                    updates.append(
                        {
                            **proposal.to_json(),
                            "status": "update_available",
                            "update_available": True,
                            "diff": diff,
                        }
                    )
                    continue
                if not valid_aoi_name(name):
                    skipped.append(
                        {"file": result.path.name, "reason": f"invalid AOI name {name!r}"}
                    )
                    continue
                try:
                    check_collision(name, template_path=state.template_path, custom_dir=None)
                except ValueError:
                    skipped.append(
                        {"file": result.path.name, "reason": "collides with a template AOI"}
                    )
                    continue
                proposal = register_proposal(
                    state, build_proposal(result.definition, result.path.name, result.path)
                )
                proposals.append(proposal.to_json())
            return {"proposals": proposals, "skipped": skipped, "updates": updates}

    @app.post("/api/aois/import/confirm")
    def confirm_import(body: ConfirmRequest):
        """Confirm a proposal: it becomes a custom AOI tab.

        Columns are re-derived from the stored parsed proposal — the
        client's ``selected`` list only filters them. The proposal is
        consumed on success: confirming the same id again is a 404.
        409 when the AOI name collides with a template AOI. When the
        AOI is already imported, the confirm is an in-place update
        (all-or-nothing, ``selected`` ignored) and refuses with 409
        while staged changes exist on that AOI. No connection needed;
        nothing here touches the PLC.
        """
        with state.lock:
            proposal = state.proposals.get(body.proposal_id)
            if proposal is None:
                raise HTTPException(
                    status_code=404, detail=f"unknown proposal {body.proposal_id!r}"
                )
            if not valid_aoi_name(proposal.aoi):
                raise HTTPException(status_code=400, detail=f"invalid AOI name {proposal.aoi!r}")
            try:
                check_collision(proposal.aoi, template_path=state.template_path, custom_dir=None)
            except ValueError as exc:
                raise HTTPException(status_code=409, detail=str(exc))
            canonical = _canonical_custom_name(state, proposal.aoi)
            if canonical is not None:
                return _confirm_update(state, proposal, canonical)
            derived = {column.subtag: column for column in proposal.columns}
            selected = body.selected if body.selected is not None else list(derived)
            if not selected:
                raise HTTPException(status_code=400, detail="no columns selected")
            unknown = [subtag for subtag in selected if subtag not in derived]
            if unknown:
                raise HTTPException(
                    status_code=400, detail=f"unknown subtags in selection: {unknown}"
                )
            wanted = set(selected)
            columns = [column for column in proposal.columns if column.subtag in wanted]
            try:
                save_custom_aoi(
                    proposal.aoi,
                    columns,
                    proposal.source_file,
                    proposal.revision,
                    path=state.custom_aoi_dir,
                    source_sha256=_source_hash(proposal),
                )
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            state.confirmed_meta[proposal.aoi] = {column.subtag: column for column in columns}
            # Learn string-UDT sizes from the import (pinned sizes win;
            # conflicts come back as warnings, never as failures).
            warnings = merge_learned_sizes(
                proposal.learned_sizes,
                proposal.source_file,
                path=state.string_registry_path,
            )
            consume_proposal(state, proposal.proposal_id)
            state.aoi_model = None
            state.drop_aoi_data(proposal.aoi)
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "import",
                    "tags": len(columns),
                    "results": {"aoi": proposal.aoi, "columns": len(columns)},
                }
            )
            response = {"aoi": proposal.aoi, "columns": len(columns)}
            if warnings:
                response["warnings"] = warnings
            return response

    @app.delete("/api/aois/custom/{name}")
    def delete_custom_aoi(name: str):
        """Remove a custom AOI (undo for a bad import). 404 when the
        name is not an imported custom AOI. The model cache and any
        baseline/staged data for the AOI are dropped with it."""
        with state.lock:
            if not valid_aoi_name(name) or name not in _load_custom_names(state):
                raise HTTPException(status_code=404, detail=f"unknown custom AOI {name!r}")
            target = _find_custom_aoi_file(state, name)
            if target is None:
                raise HTTPException(status_code=404, detail=f"unknown custom AOI {name!r}")
            try:
                definitions = load_aoi_definitions(target)
            except ValueError as exc:
                raise HTTPException(
                    status_code=503, detail=f"custom AOI store could not be read: {exc}"
                )
            if set(definitions) - {name}:
                raise HTTPException(
                    status_code=409,
                    detail=(
                        f"custom AOI {name!r} shares {target.name} with other AOIs; "
                        "remove it from the file by hand"
                    ),
                )
            try:
                target.unlink()
            except OSError as exc:
                raise HTTPException(
                    status_code=503, detail=f"could not remove custom AOI {name!r}: {exc}"
                )
            try:
                confirmed_meta_path(target).unlink()
            except OSError:
                pass  # a missing sidecar is fine; a stuck one is harmless
            state.aoi_model = None
            state.drop_aoi_data(name)
            state.column_meta.pop(name, None)
            state.confirmed_meta.pop(name, None)
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "remove_custom_aoi",
                    "tags": len(definitions[name]),
                    "results": {"aoi": name},
                }
            )
            return {"removed": name}

    @app.get("/api/aois/{aoi}")
    def read_aoi(aoi: str):
        with state.lock:
            _require_connected(state)
            _ensure_baseline(state, aoi)
            model = state.aoi_model[aoi]
            rows = [
                {"name": instance, "values": state.baseline[aoi][instance]}
                for instance in model.instances
            ]
            return {
                "name": aoi,
                "columns": [
                    {
                        "subtag": column.subtag,
                        "data_type": column.data_type,
                        "description": column.description,
                    }
                    for column in model.columns
                ],
                "rows": rows,
            }

    @app.post("/api/aois/{aoi}/paste")
    def paste(aoi: str, body: PasteRequest):
        with state.lock:
            _require_connected(state)
            _ensure_baseline(state, aoi)
            model = state.aoi_model[aoi]
            changes, rejected = parse_tsv(
                body.tsv,
                model.instances,
                model.columns,
                anchor_instance=body.anchor_instance,
                anchor_subtag=body.anchor_subtag,
                registry_path=state.string_registry_path,
            )
            applied = 0
            for change in changes:
                instance, subtag = change["instance"], change["subtag"]
                old_value = state.baseline[aoi][instance][subtag]
                if change["value"] == old_value:
                    continue  # no-op: pasted value already matches the PLC
                key = (aoi, instance, subtag)
                state.staged[key] = {
                    "aoi": aoi,
                    "instance": instance,
                    "subtag": subtag,
                    "tag": change["tag"],
                    "old_value": old_value,
                    "new_value": change["value"],
                }
                applied += 1
            staged_here = [change for key, change in state.staged.items() if key[0] == aoi]
            return {"applied": applied, "rejected": rejected, "staged": staged_here}

    @app.get("/api/changes")
    def list_changes():
        # Snapshot under the lock: a concurrent paste/write mutating
        # the dict while list() iterates it would 500 with
        # "dictionary changed size during iteration".
        with state.lock:
            return list(state.staged.values())

    @app.delete("/api/changes")
    def discard_changes():
        """Discard all staged changes. The PLC baseline is untouched."""
        with state.lock:
            discarded = len(state.staged)
            state.staged.clear()
            return {"discarded": discarded}

    @app.post("/api/write")
    def write(body: WriteRequest):
        with state.lock:
            _require_connected(state)
            # Double-check 1: the caller must name the PLC it thinks it is
            # writing to, and it must be the PLC actually connected.
            if body.plc_name != state.plc_name:
                raise HTTPException(
                    status_code=409,
                    detail=(
                        f"PLC name mismatch: connected to {state.plc_name!r}, "
                        f"write requested for {body.plc_name!r}"
                    ),
                )
            # Double-check 2: a profile's expected name must also match.
            if (
                state.profile is not None
                and state.profile.expected_plc_name is not None
                and state.profile.expected_plc_name != state.plc_name
            ):
                raise HTTPException(
                    status_code=409,
                    detail=(
                        "PLC name mismatch: profile "
                        f"{state.profile.name!r} expects "
                        f"{state.profile.expected_plc_name!r} but connected to "
                        f"{state.plc_name!r}"
                    ),
                )
            if not state.staged:
                return {"results": [], "summary": {"verified": 0, "mismatch": 0, "failed": 0}}
            # The model may have been invalidated since the changes were
            # staged (e.g. a custom AOI was imported or removed, which
            # drops the cached model) — rebuild it here rather than
            # crashing on a None model in the diff loop below.
            _ensure_model(state)

            # Auto-snapshot: capture the pre-write baseline first, so
            # this write can be rolled back from its own snapshot. A
            # failed capture never blocks the write — it is logged,
            # recorded in the audit entry, and reported in the
            # response as auto_snapshot {ok: false, error}.
            auto_snapshot = _take_auto_snapshot(state)

            # Re-read the STAGED tags only and diff with the core library,
            # so only staged differences are written. A staged tag absent
            # from the diffs is already at the staged value in the PLC, so
            # it verifies without a write. Unstaged tags are never
            # compared: if a PLC value drifted from the baseline since the
            # tab was loaded, writing the stale baseline back over it
            # would silently revert a change the user never made.
            diffs: list[tuple[str, object]] = []
            by_instance: dict[tuple[str, str], list[tuple[str, str]]] = {}
            for key in state.staged:
                by_instance.setdefault((key[0], key[1]), []).append(key)
            try:
                for (aoi, instance), keys in by_instance.items():
                    staged_subtags = [key[2] for key in keys]
                    sheet_row = [
                        (
                            instance + sub_tag,
                            state.staged[(aoi, instance, sub_tag)]["new_value"],
                        )
                        for sub_tag in staged_subtags
                    ]
                    tag_list = make_tag_list(instance, staged_subtags)
                    plc_row, _raw = read_plc_row(state.plc, tag_list)
                    diffs.extend(row_differences(sheet_row, plc_row))
            except HTTPException:
                raise
            except Exception as exc:
                raise HTTPException(
                    status_code=502, detail=f"failed to re-read PLC before write: {exc}"
                )

            staged_by_tag = {change["tag"]: change for change in state.staged.values()}
            intended = dict(diffs)
            status: dict[str, str] = {}
            # Staged tags the PLC already holds at the staged value verify
            # at once — no write needed (and no write to fail).
            for tag in staged_by_tag:
                if tag not in intended:
                    status[tag] = "verified"

            # Chunked sequential writes, one PLC at a time (lock is held).
            chunk_size = max(1, state.settings.write_chunk_size)
            try:
                for start in range(0, len(diffs), chunk_size):
                    chunk = diffs[start : start + chunk_size]
                    chunk_results = write_plc_row(state.plc, chunk)
                    for (tag, _value), ok in zip(chunk, chunk_results):
                        if not ok:
                            status[tag] = "failed"
            except Exception as exc:
                raise HTTPException(status_code=502, detail=f"failed to write to PLC: {exc}")

            # Read-back verification: re-read every successfully written
            # tag and compare with what was intended. A mismatch (or an
            # unreadable read-back) is re-read up to read_back_retries
            # times before it is declared a mismatch — a transient read
            # glitch must not condemn a good write.
            written_tags = [tag for tag in intended if tag not in status]
            if written_tags:
                read_back = _read_back_values(state.plc, written_tags)
                pending = [
                    tag
                    for tag in written_tags
                    if not values_equal(
                        intended[tag],
                        read_back.get(tag),
                        state.settings.real_tolerance,
                        READBACK_ABSOLUTE_TOLERANCE,
                    )
                ]
                for _attempt in range(max(0, state.settings.read_back_retries)):
                    if not pending:
                        break
                    retry = _read_back_values(state.plc, pending)
                    pending = [
                        tag
                        for tag in pending
                        if not values_equal(
                            intended[tag],
                            retry.get(tag),
                            state.settings.real_tolerance,
                            READBACK_ABSOLUTE_TOLERANCE,
                        )
                    ]
                for tag in written_tags:
                    status[tag] = "mismatch" if tag in pending else "verified"

            # Results in staged order: the changes as the user staged them.
            results = [{"tag": tag, "status": status[tag]} for tag in staged_by_tag]
            verified_tags = {tag for tag, s in status.items() if s == "verified"}
            for key in list(state.staged):
                if state.staged[key]["tag"] in verified_tags:
                    change = state.staged.pop(key)
                    state.baseline[key[0]][key[1]][key[2]] = change["new_value"]
            summary = {
                "verified": sum(1 for r in results if r["status"] == "verified"),
                "mismatch": sum(1 for r in results if r["status"] == "mismatch"),
                "failed": sum(1 for r in results if r["status"] == "failed"),
            }
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "write",
                    "tags": len(results),
                    "results": summary,
                    "auto_snapshot": auto_snapshot,
                }
            )
            return {"results": results, "summary": summary, "auto_snapshot": auto_snapshot}

    @app.get("/api/audit")
    def audit():
        with state.lock:
            return state.audit.list()

    @app.get("/api/export")
    def export():
        """Download the current snapshot as ``.xlsx`` in template layout.

        Baseline values for every AOI plus staged changes on top (amber
        fill). The file name carries the PLC name — the CLI write mode
        refuses a file whose name does not contain it. When disconnected
        there is no baseline, so the template is exported as-is under
        the fallback name.
        """
        with state.lock:
            model = _ensure_model(state)
            if state.connected:
                for aoi in model:
                    _ensure_baseline(state, aoi)
            book = build_export_workbook(
                state.template_path,
                state.baseline,
                list(state.staged.values()),
                state.plc_name,
            )
            buf = BytesIO()
            book.save(buf)
            book.close()
            buf.seek(0)
            return StreamingResponse(
                buf,
                media_type=("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                headers={
                    "Content-Disposition": (
                        f'attachment; filename="{_export_filename(state.plc_name)}"'
                    )
                },
            )

    # -- PLC snapshots (Phase 9a) -------------------------------------

    @app.post("/api/snapshots", status_code=201)
    def create_snapshot(body: SnapshotRequest):
        """Save the live baseline for every AOI as a named snapshot.

        201 ``{name, plc_name, created_at, aoi_count,
        instance_count}``. A manual ``name`` must be
        ``[A-Za-z0-9_-]+`` (400 otherwise) and must not already exist
        (409 — a named snapshot is never silently overwritten). With
        no name, the snapshot is saved as
        ``<plc>_<YYYYMMDD-HHMMSS>.json``.
        """
        with state.lock:
            _require_connected(state)
            try:
                name = sanitize_snapshot_name(body.name) if body.name is not None else None
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            _refresh_all_baselines(state)
            snapshot = build_snapshot(state.plc_name, state.baseline, _aoi_version(state))
            try:
                saved = save_snapshot(state.snapshot_dir, snapshot, name=name)
            except ValueError as exc:
                raise HTTPException(status_code=409, detail=str(exc))
            except OSError as exc:
                raise HTTPException(status_code=503, detail=f"snapshot could not be saved: {exc}")
            aoi_count, instance_count = snapshot_counts(snapshot)
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "snapshot",
                    "tags": sum(
                        len(values)
                        for instances in snapshot["aois"].values()
                        for values in instances.values()
                    ),
                    "results": {"name": saved.stem, "auto": False},
                }
            )
            return {
                "name": saved.stem,
                "plc_name": snapshot["plc_name"],
                "created_at": snapshot["created_at"],
                "aoi_count": aoi_count,
                "instance_count": instance_count,
            }

    @app.get("/api/snapshots")
    def list_snapshots_route():
        """Snapshot metadata, newest first, with an ``auto`` flag.
        No connection needed — it is a file listing."""
        with state.lock:
            return list_snapshots(state.snapshot_dir)

    @app.get("/api/snapshots/{name}")
    def download_snapshot(name: str):
        """Download one snapshot file (JSON). 404 when unknown.

        The raw bytes are served without parsing: a corrupt snapshot
        is still downloadable for inspection (compare/restore refuse
        it, delete removes it).
        """
        with state.lock:
            path = _resolve_snapshot_file(state, name)
            try:
                content = path.read_bytes()
            except OSError as exc:
                raise HTTPException(
                    status_code=503, detail=f"snapshot {name!r} could not be read: {exc}"
                )
            return Response(
                content=content,
                media_type="application/json",
                headers={"Content-Disposition": f'attachment; filename="{path.name}"'},
            )

    @app.post("/api/snapshots/upload", status_code=201)
    async def upload_snapshot(request: Request):
        """Upload a snapshot file (``.json`` only, 5 MB cap).

        The file is validated before it is stored (never written to
        disk first). 409 when the file's ``plc_name`` differs from the
        connected PLC — a snapshot for another PLC must never be
        restored here by accident. 201 on success, with the stored
        name (the uploaded filename stem).
        """
        body = await request.body()
        content_type = request.headers.get("content-type", "")
        with state.lock:
            _require_connected(state)
            try:
                filename, data = parse_multipart_upload(body, content_type)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            if Path(filename).suffix.lower() != ".json":
                raise HTTPException(
                    status_code=400,
                    detail=f"only .json snapshot files can be uploaded, got {filename!r}",
                )
            try:
                snapshot = parse_snapshot_bytes(data)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            if snapshot["plc_name"] != state.plc_name:
                raise HTTPException(
                    status_code=409,
                    detail=(
                        f"snapshot is for PLC {snapshot['plc_name']!r}, "
                        f"connected to {state.plc_name!r}"
                    ),
                )
            try:
                name = sanitize_snapshot_name(Path(filename).stem)
            except ValueError as exc:
                raise HTTPException(status_code=400, detail=str(exc))
            try:
                saved = save_snapshot(state.snapshot_dir, snapshot, name=name)
            except ValueError as exc:
                raise HTTPException(status_code=409, detail=str(exc))
            except OSError as exc:
                raise HTTPException(status_code=503, detail=f"snapshot could not be saved: {exc}")
            aoi_count, instance_count = snapshot_counts(snapshot)
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "snapshot",
                    "tags": sum(
                        len(values)
                        for instances in snapshot["aois"].values()
                        for values in instances.values()
                    ),
                    "results": {"name": saved.stem, "auto": False, "uploaded": True},
                }
            )
            return {
                "name": saved.stem,
                "plc_name": snapshot["plc_name"],
                "created_at": snapshot["created_at"],
                "aoi_count": aoi_count,
                "instance_count": instance_count,
            }

    @app.post("/api/snapshots/{name}/compare")
    def compare_snapshot(name: str):
        """Diff a snapshot against the live PLC (freshly re-read).

        ``{plc_name, differences: [{aoi, instance, subtag,
        snapshot_value, current_value}]}``. ``plc_name`` is the PLC
        the snapshot was taken from, so the UI can warn when it is not
        the connected PLC (the diff is still computed — comparing
        across PLCs is a legitimate use). 409 when not connected, 404
        for an unknown snapshot.
        """
        with state.lock:
            _require_connected(state)
            snapshot, _path = _load_named_snapshot(state, name)
            for aoi in snapshot["aois"]:
                if aoi in _ensure_model(state):
                    _refresh_baseline(state, aoi)
            return {
                "plc_name": snapshot["plc_name"],
                "differences": diff_snapshot(snapshot["aois"], state.baseline),
            }

    @app.post("/api/snapshots/{name}/restore")
    def restore_snapshot(name: str):
        """Stage a snapshot's differing values — nothing is written.

        Values go through the normal validation path; failures are
        reported as ``skipped`` with reasons and are never staged.
        409 when the snapshot is for another PLC or when not
        connected, 404 for an unknown snapshot.
        """
        with state.lock:
            _require_connected(state)
            snapshot, _path = _load_named_snapshot(state, name)
            if snapshot["plc_name"] != state.plc_name:
                raise HTTPException(
                    status_code=409,
                    detail=(
                        f"snapshot is for PLC {snapshot['plc_name']!r}, "
                        f"connected to {state.plc_name!r}"
                    ),
                )
            staged, skipped = _restore_from_snapshot(state, snapshot)
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "snapshot_restore",
                    "tags": staged,
                    "results": {"name": name, "staged": staged, "skipped": len(skipped)},
                }
            )
            return {"staged": staged, "skipped": skipped}

    @app.delete("/api/snapshots/{name}")
    def delete_snapshot(name: str):
        """Delete a snapshot file. 404 when unknown.

        The file is removed without parsing it: a corrupt snapshot
        must be deletable, or it would be stuck in the store forever.
        """
        with state.lock:
            path = _resolve_snapshot_file(state, name)
            try:
                path.unlink()
            except OSError as exc:
                raise HTTPException(
                    status_code=503, detail=f"snapshot {name!r} could not be removed: {exc}"
                )
            state.audit.record(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "snapshot_delete",
                    "tags": 0,
                    "results": {"name": name},
                }
            )
            return {"removed": name}

    static_dir = Path(__file__).resolve().parent / "static"
    if static_dir.is_dir():
        app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
    else:

        @app.get("/", response_class=HTMLResponse)
        def index():
            return _PLACEHOLDER_PAGE

    return app


def main():
    """Entry point: ``python -m pypax.web`` — localhost only, never 0.0.0.0."""
    import uvicorn

    settings, profiles = load_app_config()
    app = create_app(settings=settings, profiles=profiles)
    uvicorn.run(app, host=settings.host, port=settings.port)


if __name__ == "__main__":
    main()
