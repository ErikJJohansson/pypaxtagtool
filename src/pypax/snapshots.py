"""PLC snapshots (Phase 9a): save, load, list, diff, and restore baselines.

A snapshot is a JSON file holding the PLC baseline for every AOI
(template + custom) at one moment: ``{"format": "pypax-snapshot/1",
"plc_name", "created_at" (ISO 8601 with offset), "aoi_version",
"aois": {AOI: {instance: {subtag: value}}}}``. Values are JSON-native
(REAL float, BOOL 0/1 int, strings str) — the same formatted values
:func:`pypax.plc.read_plc_row` returns, so a snapshot round-trips
through the normal staging path without re-formatting.

Storage: ``snapshots/`` (gitignored), auto-snapshots under
``snapshots/auto/``. Manual names are sanitized to ``[A-Za-z0-9_-]+``
(no path separators, so a name can never escape the directory).
Auto-snapshots are named ``<plc>_<YYYYMMDD-HHMMSS>.json``.

Files are treated as untrusted input on load: a snapshot is only
accepted when its shape, format tag, and value types check out —
anything else raises ``ValueError`` and is never staged or written.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

FORMAT = "pypax-snapshot/1"
MAX_UPLOAD_BYTES = 5 * 1024 * 1024  # 5 MB cap, same as the .L5X upload
AUTO_KEEP_PER_PLC = 10  # newest auto-snapshots kept per PLC

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SNAPSHOT_DIR_NAME = "snapshots"
AUTO_DIR_NAME = "auto"

_NAME_RE = re.compile(r"^[A-Za-z0-9_-]+$")


def resolve_snapshot_dir(settings, explicit: str | Path | None = None) -> Path:
    """Snapshot directory: explicit path, else the setting, else ``snapshots/``.

    Relative paths resolve against the repo root (the task's default);
    absolute paths are used as-is.
    """
    if explicit is not None:
        return Path(explicit)
    raw = (getattr(settings, "snapshot_dir", None) if settings is not None else None) or (
        DEFAULT_SNAPSHOT_DIR_NAME
    )
    path = Path(raw)
    return path if path.is_absolute() else REPO_ROOT / path


def _normalize_value(value):
    """One snapshot value, JSON-native: bool -> 0/1 int (the BOOL
    convention), int/float/str/None pass through; anything else (a
    list, dict, or exotic type) is not a PLC scalar and is rejected."""
    if isinstance(value, bool):
        return int(value)
    if value is None or isinstance(value, (int, float, str)):
        return value
    raise ValueError(f"snapshot value {value!r} is not a JSON scalar")


def build_snapshot(plc_name, aois, aoi_version, created_at=None) -> dict:
    """Build a snapshot dict from baseline values.

    ``aois`` is ``{AOI: {instance: {subtag: value}}}``; values are
    normalized (bool -> 0/1) and the result must be JSON-serializable
    or ``ValueError`` is raised — a snapshot that cannot be written
    is a bug, not something to discover at restore time.
    """
    if not isinstance(plc_name, str) or not plc_name:
        raise ValueError("snapshot needs a PLC name")
    if not isinstance(aoi_version, str) or not aoi_version:
        raise ValueError("snapshot needs an AOI version")
    if not isinstance(aois, dict):
        raise ValueError("snapshot aois must be a dict")
    stamp = created_at or datetime.now(timezone.utc)
    if stamp.tzinfo is None:
        raise ValueError("snapshot created_at must carry a timezone offset")
    normalized = {}
    for aoi, instances in aois.items():
        if not isinstance(instances, dict):
            raise ValueError(f"snapshot AOI {aoi!r} is not a dict of instances")
        normalized[aoi] = {}
        for instance, values in instances.items():
            if not isinstance(values, dict):
                raise ValueError(f"snapshot instance {instance!r} is not a dict of values")
            normalized[aoi][instance] = {
                subtag: _normalize_value(value) for subtag, value in values.items()
            }
    snapshot = {
        "format": FORMAT,
        "plc_name": plc_name,
        "created_at": stamp.isoformat(),
        "aoi_version": aoi_version,
        "aois": normalized,
    }
    try:
        json.dumps(snapshot)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"snapshot is not JSON-serializable: {exc}") from exc
    return snapshot


def snapshot_counts(snapshot: dict) -> tuple[int, int]:
    """``(aoi_count, instance_count)`` for a snapshot dict."""
    aois = snapshot.get("aois", {})
    return len(aois), sum(len(instances) for instances in aois.values())


def sanitize_snapshot_name(name) -> str:
    """Validate a manual snapshot name: ``[A-Za-z0-9_-]+`` only.

    No path separators, no dots, no spaces — a name is a filename
    stem and must never escape the snapshot directory.
    """
    if not isinstance(name, str) or not _NAME_RE.match(name):
        raise ValueError(
            "snapshot name must be one or more of [A-Za-z0-9_-] "
            f"(no spaces, dots, or path separators), got {name!r}"
        )
    return name


def snapshot_filename(plc_name: str, created_at: datetime) -> str:
    """Auto-snapshot filename stem: ``<plc>_<YYYYMMDD-HHMMSS>``.

    The PLC name is sanitized the same way as a manual name (anything
    outside ``[A-Za-z0-9_-]`` becomes ``_``) so the stem is always a
    safe filename.
    """
    safe = "".join(c if c.isalnum() or c in "_-" else "_" for c in plc_name)
    return f"{safe}_{created_at.strftime('%Y%m%d-%H%M%S')}"


def _auto_dir(snapshot_dir: str | Path) -> Path:
    return Path(snapshot_dir) / AUTO_DIR_NAME


def _resolve_file(snapshot_dir: str | Path, name: str) -> Path:
    """The file for a snapshot name: the manual file if it exists,
    else the auto file. ``name`` is validated first, so ``..`` can
    never walk out of the directory."""
    stem = sanitize_snapshot_name(name)
    manual = Path(snapshot_dir) / f"{stem}.json"
    if manual.is_file():
        return manual
    return _auto_dir(snapshot_dir) / f"{stem}.json"


def save_snapshot(snapshot_dir, snapshot: dict, name: str | None = None) -> Path:
    """Persist a snapshot dict.

    ``name`` is a manual name (validated); ``None`` means an
    auto-snapshot, named ``<plc>_<YYYYMMDD-HHMMSS>`` with a ``-2``,
    ``-3`` suffix when two land in the same second. A manual name
    that already exists is refused (``ValueError``) — overwriting a
    named snapshot would silently destroy a safety artifact. Writes
    via a temp file + rename so a crash cannot leave a half-written
    snapshot behind. Returns the saved path.
    """
    directory = Path(snapshot_dir)
    if name is not None:
        stem = sanitize_snapshot_name(name)
        directory.mkdir(parents=True, exist_ok=True)
        target = directory / f"{stem}.json"
        if target.exists() or (_auto_dir(directory) / f"{stem}.json").exists():
            # Lookup by name is single-valued (manual wins), so a stem
            # shared with an auto-snapshot would leave one of the two
            # unreachable by name.
            raise ValueError(f"snapshot {name!r} already exists")
    else:
        stamp = datetime.fromisoformat(snapshot["created_at"])
        stem = snapshot_filename(snapshot["plc_name"], stamp)
        auto = _auto_dir(directory)
        auto.mkdir(parents=True, exist_ok=True)
        target = auto / f"{stem}.json"
        suffix = 2
        while target.exists() or (directory / target.name).exists():
            target = auto / f"{stem}-{suffix}.json"
            suffix += 1
    payload = json.dumps(snapshot, indent=2, sort_keys=True).encode("utf-8")
    fd, tmp_path = tempfile.mkstemp(dir=str(target.parent), suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(payload)
        os.replace(tmp_path, target)
    except BaseException:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise
    return target


def _validate_snapshot(data) -> dict:
    """Check a parsed JSON value is a snapshot, and normalize it.

    Raises ``ValueError`` for anything that is not a well-formed
    ``pypax-snapshot/1`` document — wrong format tag, wrong shapes,
    non-scalar values. Bool leaves become 0/1 ints.
    """
    if not isinstance(data, dict):
        raise ValueError("snapshot must be a JSON object")
    if data.get("format") != FORMAT:
        raise ValueError(f"not a {FORMAT} snapshot (format is {data.get('format')!r})")
    for key in ("plc_name", "created_at", "aoi_version"):
        if not isinstance(data.get(key), str) or not data[key]:
            raise ValueError(f"snapshot field {key!r} must be a non-empty string")
    try:
        stamp = datetime.fromisoformat(data["created_at"])
    except ValueError as exc:
        raise ValueError(f"snapshot created_at is not ISO 8601: {data['created_at']!r}") from exc
    if stamp.tzinfo is None:
        raise ValueError("snapshot created_at must carry a timezone offset")
    aois = data.get("aois")
    if not isinstance(aois, dict):
        raise ValueError("snapshot field 'aois' must be an object")
    for aoi, instances in aois.items():
        if not isinstance(instances, dict):
            raise ValueError(f"snapshot AOI {aoi!r} is not an object")
        for instance, values in instances.items():
            if not isinstance(values, dict):
                raise ValueError(f"snapshot instance {instance!r} is not an object")
            for subtag, value in values.items():
                values[subtag] = _normalize_value(value)
    return data


def load_snapshot(path: str | Path) -> dict:
    """Load and validate a snapshot file. Raises ``ValueError`` for
    corrupt JSON or a malformed snapshot, ``FileNotFoundError`` when
    the file does not exist."""
    try:
        raw = Path(path).read_bytes()
    except OSError as exc:
        if isinstance(exc, FileNotFoundError):
            raise
        raise ValueError(f"cannot read snapshot {path}: {exc}") from exc
    try:
        data = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError(f"snapshot {path} is not valid JSON: {exc}") from exc
    return _validate_snapshot(data)


def parse_snapshot_bytes(data: bytes) -> dict:
    """Validate uploaded snapshot bytes (never written to disk first)."""
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(f"snapshot exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB cap")
    try:
        parsed = json.loads(data.decode("utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise ValueError(f"snapshot is not valid JSON: {exc}") from exc
    return _validate_snapshot(parsed)


def find_snapshot_file(snapshot_dir: str | Path, name: str) -> Path | None:
    """The file for a snapshot name (manual, else auto), or ``None``
    when no snapshot with that name exists. ``name`` is validated
    first, so ``..`` can never walk out of the directory."""
    path = _resolve_file(snapshot_dir, name)
    return path if path.is_file() else None


def _snapshot_meta(path: Path, auto: bool) -> dict | None:
    """List-entry metadata for one snapshot file, or ``None`` when the
    file cannot be read as a snapshot (a corrupt file must not break
    the listing — it stays addressable by name for download/delete)."""
    try:
        snapshot = load_snapshot(path)
    except (ValueError, FileNotFoundError, OSError):
        return None
    aoi_count, instance_count = snapshot_counts(snapshot)
    return {
        "name": path.stem,
        "plc_name": snapshot["plc_name"],
        "created_at": snapshot["created_at"],
        "aoi_count": aoi_count,
        "instance_count": instance_count,
        "auto": auto,
    }


def list_snapshots(snapshot_dir: str | Path) -> list[dict]:
    """Snapshot metadata, newest first (by created_at, then name).

    Manual snapshots first by file, then auto-snapshots; the combined
    list is sorted by created_at descending. Unreadable files are
    skipped. A name shared by a manual and an auto file (only possible
    for hand-placed files — ``save_snapshot`` refuses the collision)
    lists once, the manual one: name lookup resolves manual-first,
    so the list must match what a name actually addresses.
    """
    directory = Path(snapshot_dir)
    entries: list[dict] = []
    if directory.is_dir():
        for path in sorted(directory.glob("*.json")):
            meta = _snapshot_meta(path, auto=False)
            if meta is not None:
                entries.append(meta)
    auto = _auto_dir(directory)
    if auto.is_dir():
        for path in sorted(auto.glob("*.json")):
            meta = _snapshot_meta(path, auto=True)
            if meta is not None:
                entries.append(meta)
    seen: set[str] = set()
    deduped: list[dict] = []
    for entry in entries:
        if entry["name"] in seen:
            continue
        seen.add(entry["name"])
        deduped.append(entry)
    deduped.sort(key=lambda entry: (entry["created_at"], entry["name"]), reverse=True)
    return deduped


def diff_snapshot(snapshot_aois: dict, current_aois: dict) -> list[dict]:
    """Differences between a snapshot and the current baseline.

    One entry per (aoi, instance, subtag) that differs, sorted by
    (aoi, instance, subtag): a changed value, a tag only in the
    snapshot (``current_value`` is ``None`` — the tag is gone from the
    PLC), or a tag only in the current baseline (``snapshot_value``
    is ``None`` — it appeared after the snapshot). Equal values are
    not differences.
    """
    differences: list[dict] = []
    aois = set(snapshot_aois) | set(current_aois)
    for aoi in sorted(aois):
        snap_instances = snapshot_aois.get(aoi, {})
        current_instances = current_aois.get(aoi, {})
        for instance in sorted(set(snap_instances) | set(current_instances)):
            snap_values = snap_instances.get(instance, {})
            current_values = current_instances.get(instance, {})
            for subtag in sorted(set(snap_values) | set(current_values)):
                snap_value = snap_values.get(subtag)
                current_value = current_values.get(subtag)
                if snap_value != current_value:
                    differences.append(
                        {
                            "aoi": aoi,
                            "instance": instance,
                            "subtag": subtag,
                            "snapshot_value": snap_value,
                            "current_value": current_value,
                        }
                    )
    return differences


def prune_auto_snapshots(snapshot_dir, plc_name: str, keep: int = AUTO_KEEP_PER_PLC) -> int:
    """Keep the newest ``keep`` auto-snapshots for ``plc_name``, delete
    older ones. Files that cannot be read as snapshots are never
    deleted (and do not count towards ``keep``). Returns how many
    files were removed."""
    auto = _auto_dir(snapshot_dir)
    if not auto.is_dir():
        return 0
    stamped: list[tuple[str, Path]] = []
    for path in sorted(auto.glob("*.json")):
        try:
            snapshot = load_snapshot(path)
        except (ValueError, FileNotFoundError, OSError):
            continue
        if snapshot["plc_name"] == plc_name:
            stamped.append((snapshot["created_at"], path))
    stamped.sort(key=lambda item: (item[0], item[1].name))
    removed = 0
    for _stamp, path in stamped[: max(0, len(stamped) - keep)]:
        try:
            path.unlink()
            removed += 1
        except OSError:
            pass
    return removed
