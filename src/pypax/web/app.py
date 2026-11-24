"""FastAPI backend for the PyPaxTagTool webapp.

Implements the frozen REST contract (see the Phase 3 brief): connect,
PLC status, AOI list/read, TSV paste, staged changes, guarded writes,
audit log, and a not-yet-implemented export. All state lives in one
:class:`AppState` guarded by a lock; one PLC connection at a time.

The backend reads through the same ``pypax`` core library as the CLI:
instances are discovered from the PLC, columns come from the template
workbook, reads use :func:`pypax.plc.read_plc_row`, and writes use the
ordered :func:`pypax.diff.row_differences` plus
:func:`pypax.plc.write_plc_row`. Errors are JSON ``{"detail": ...}``;
tracebacks never leak.
"""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from pycomm3 import LogixDriver

from ..config import AppSettings, load_app_config
from ..diff import row_differences
from ..plc import make_tag_list, read_plc_row, write_plc_row
from .aoi_model import discover_instances, load_template_model, read_aoi_rows
from .paste import parse_tsv
from .state import AppState

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


def default_plc_factory(commpath):
    """Build a real pycomm3 driver for a comm path."""
    return LogixDriver(commpath, init_tags=True, init_program_tags=True)


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
    """Template AOI model, cached per connection; 503 when unavailable."""
    if state.aoi_model is None:
        if state.template_path is None:
            raise HTTPException(
                status_code=503,
                detail="template workbook not found: set template_path or place "
                f"{DEFAULT_TEMPLATE_NAME} in the working directory",
            )
        try:
            state.aoi_model = load_template_model(state.template_path)
        except FileNotFoundError:
            raise HTTPException(
                status_code=503,
                detail=f"template workbook not found: {state.template_path}",
            )
        except Exception as exc:  # unreadable workbook: report, don't leak
            raise HTTPException(
                status_code=503, detail=f"template workbook could not be read: {exc}"
            )
    return state.aoi_model


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


def create_app(plc_factory=None, template_path=None, settings=None, profiles=None):
    """Build the FastAPI app. Tests inject a fake PLC via ``plc_factory``."""
    settings = settings or AppSettings()
    profiles = profiles or {}
    state = AppState(
        plc_factory=plc_factory or default_plc_factory,
        template_path=_resolve_template_path(settings, template_path),
        settings=settings,
        profiles=profiles,
    )
    app = FastAPI(title="PyPaxTagTool")
    app.state.pypax = state  # tests and tooling inspect the live state here

    @app.get("/api/health")
    def health():
        return {"status": "ok"}

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
            return {"connected": True, "plc_name": plc_name, "commpath": body.commpath}

    @app.post("/api/disconnect")
    def disconnect():
        with state.lock:
            if state.connected:
                try:
                    state.plc.close()
                except Exception:
                    pass
            state.reset_connection()
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
                    }
                )
            return result

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
        return list(state.staged.values())

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
                return {"results": [], "summary": {"written": 0, "failed": 0}}

            # Re-read the affected rows and diff with the core library, so
            # only real differences are written (and only staged ones).
            diffs: list[tuple[str, object]] = []
            by_instance: dict[tuple[str, str], list[tuple[str, str]]] = {}
            for key in state.staged:
                by_instance.setdefault((key[0], key[1]), []).append(key)
            try:
                for (aoi, instance), keys in by_instance.items():
                    model = state.aoi_model[aoi]
                    sub_tags = [column.subtag for column in model.columns]
                    sheet_row = [
                        (
                            instance + sub_tag,
                            state.staged[(aoi, instance, sub_tag)]["new_value"]
                            if (aoi, instance, sub_tag) in state.staged
                            else state.baseline[aoi][instance][sub_tag],
                        )
                        for sub_tag in sub_tags
                    ]
                    tag_list = make_tag_list(instance, sub_tags)
                    plc_row, _raw = read_plc_row(state.plc, tag_list)
                    diffs.extend(row_differences(sheet_row, plc_row))
            except HTTPException:
                raise
            except Exception as exc:
                raise HTTPException(
                    status_code=502, detail=f"failed to re-read PLC before write: {exc}"
                )

            # Chunked sequential writes, one PLC at a time (lock is held).
            results: list[dict] = []
            chunk_size = max(1, state.settings.write_chunk_size)
            try:
                for start in range(0, len(diffs), chunk_size):
                    chunk = diffs[start : start + chunk_size]
                    chunk_results = write_plc_row(state.plc, chunk)
                    for (tag, _value), ok in zip(chunk, chunk_results):
                        results.append({"tag": tag, "status": "written" if ok else "failed"})
            except Exception as exc:
                raise HTTPException(status_code=502, detail=f"failed to write to PLC: {exc}")

            written_tags = {r["tag"] for r in results if r["status"] == "written"}
            for key in list(state.staged):
                if state.staged[key]["tag"] in written_tags:
                    change = state.staged.pop(key)
                    state.baseline[key[0]][key[1]][key[2]] = change["new_value"]
            summary = {
                "written": sum(1 for r in results if r["status"] == "written"),
                "failed": sum(1 for r in results if r["status"] == "failed"),
            }
            state.audit.append(
                {
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "plc_name": state.plc_name,
                    "action": "write",
                    "tags": len(diffs),
                    "results": summary,
                }
            )
            return {"results": results, "summary": summary}

    @app.get("/api/audit")
    def audit():
        return list(state.audit)

    @app.get("/api/export")
    def export():
        raise HTTPException(status_code=501, detail="not implemented yet")

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
