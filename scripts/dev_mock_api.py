"""Development mock API server for the PyPaxTagTool web UI.

Serves the static frontend (``src/pypax/web/static/``) and answers the
frozen REST contract endpoints with canned in-memory data — no FastAPI, no
PLC, no network beyond localhost. Standard library only.

Run from the repo root::

    python scripts/dev_mock_api.py [--port 8000]

then open http://127.0.0.1:8000/ (or http://127.0.0.1:8000/?mock=1 is not
needed here — this server *is* the mock backend; ?mock=1 is for the
in-browser MockApi when no server is running at all).

The canned data mirrors the ``MockApi`` object in ``app.js``: two AOIs
(P_AIn with 3 instances, P_DIn with 2), PLC name ``MOCK_PLC``.
"""

from __future__ import annotations

import argparse
import copy
import json
import mimetypes
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

STATIC_DIR = Path(__file__).resolve().parents[1] / "src" / "pypax" / "web" / "static"

PLC_NAME = "MOCK_PLC"
DEFAULT_COMMPATH = "10.10.17.10/4"

AOIS = [
    {"name": "P_AIn", "instance_count": 3, "subtag_count": 4},
    {"name": "P_DIn", "instance_count": 2, "subtag_count": 3},
]

# PLC baseline values. GET /api/aois/{aoi} returns these; staged changes are
# overlaid by the frontend and only land here on a successful write.
PLC_DATA = {
    "P_AIn": {
        "name": "P_AIn",
        "columns": [
            {
                "subtag": ".Cfg_Desc",
                "data_type": "STRING_40",
                "description": "Description for display on HMI",
            },
            {"subtag": ".Cfg_HiLim", "data_type": "REAL", "description": "High alarm limit"},
            {"subtag": ".Cfg_LoLim", "data_type": "REAL", "description": "Low alarm limit"},
            {"subtag": ".Cfg_HiHiLim", "data_type": "REAL", "description": "High-high alarm limit"},
        ],
        "rows": [
            {
                "name": "FIT-101",
                "values": {
                    ".Cfg_Desc": "Flow transmitter",
                    ".Cfg_HiLim": 90,
                    ".Cfg_LoLim": 10,
                    ".Cfg_HiHiLim": 95,
                },
            },
            {
                "name": "PIT-201",
                "values": {
                    ".Cfg_Desc": "Pressure transmitter",
                    ".Cfg_HiLim": 150,
                    ".Cfg_LoLim": 5,
                    ".Cfg_HiHiLim": 160,
                },
            },
            {
                "name": "TIT-301",
                "values": {
                    ".Cfg_Desc": "Temperature transmitter",
                    ".Cfg_HiLim": 120,
                    ".Cfg_LoLim": 0,
                    ".Cfg_HiHiLim": 130,
                },
            },
        ],
    },
    "P_DIn": {
        "name": "P_DIn",
        "columns": [
            {
                "subtag": ".Cfg_Desc",
                "data_type": "STRING_40",
                "description": "Description for display on HMI",
            },
            {"subtag": ".Cfg_0StText", "data_type": "STRING_20", "description": "State 0 text"},
            {"subtag": ".Cfg_1StText", "data_type": "STRING_20", "description": "State 1 text"},
        ],
        "rows": [
            {
                "name": "XIC-301",
                "values": {
                    ".Cfg_Desc": "Valve closed switch",
                    ".Cfg_0StText": "Closed",
                    ".Cfg_1StText": "Open",
                },
            },
            {
                "name": "XIC-302",
                "values": {
                    ".Cfg_Desc": "Valve open switch",
                    ".Cfg_0StText": "Closed",
                    ".Cfg_1StText": "Open",
                },
            },
        ],
    },
}


def parse_value(raw: str, data_type: str) -> tuple[bool, object, str]:
    """Validate a pasted cell against its data type.

    Returns (ok, value, reason) mirroring the frontend's parseValue.
    """
    text = str(raw).strip()
    dtype = (data_type or "").upper()
    if "BOOL" in dtype:
        if text in ("1", "true", "True", "TRUE"):
            return True, 1, ""
        if text in ("0", "false", "False", "FALSE"):
            return True, 0, ""
        return False, None, "expected 0 or 1 for BOOL"
    if "INT" in dtype:
        try:
            return True, int(text), ""
        except ValueError:
            return False, None, "expected an integer for " + data_type
    if "REAL" in dtype:
        try:
            value = float(text)
        except ValueError:
            return False, None, "expected a number for REAL"
        if value != value or value in (float("inf"), float("-inf")):
            return False, None, "expected a number for REAL"
        return True, value, ""
    return True, text, ""


class MockState:
    """In-memory mock PLC + staged changes + audit log."""

    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.connected = False
        self.commpath = DEFAULT_COMMPATH
        self.plc_data = copy.deepcopy(PLC_DATA)
        self.staged: list[dict] = []
        self.audit: list[dict] = []

    def staged_value(self, aoi: str, instance: str, subtag: str) -> object:
        for entry in self.staged:
            if entry["aoi"] == aoi and entry["instance"] == instance and entry["subtag"] == subtag:
                return entry["new_value"]
        for row in self.plc_data[aoi]["rows"]:
            if row["name"] == instance:
                return row["values"][subtag]
        return None

    def upsert_staged(
        self, aoi: str, instance: str, subtag: str, old_value: object, new_value: object
    ) -> None:
        for entry in self.staged:
            if entry["aoi"] == aoi and entry["instance"] == instance and entry["subtag"] == subtag:
                entry["new_value"] = new_value
                return
        self.staged.append(
            {
                "aoi": aoi,
                "instance": instance,
                "subtag": subtag,
                "tag": instance + subtag,
                "old_value": old_value,
                "new_value": new_value,
            }
        )

    def apply_paste(
        self, aoi: str, tsv: str, anchor_instance: str | None, anchor_subtag: str | None
    ) -> dict:
        """Apply a pasted TSV block starting at the anchor cell.

        Returns {"applied": [...], "rejected": [...], "staged": n} per the contract.
        """
        data = self.plc_data[aoi]
        columns = [c["subtag"] for c in data["columns"]]
        rows = data["rows"]
        start_row = next((i for i, r in enumerate(rows) if r["name"] == anchor_instance), 0)
        start_col = columns.index(anchor_subtag) if anchor_subtag in columns else 0
        applied: list[dict] = []
        rejected: list[dict] = []
        for line_idx, line in enumerate(str(tsv).splitlines()):
            if line == "":
                continue
            for cell_idx, cell in enumerate(line.split("\t")):
                row = rows[start_row + line_idx] if start_row + line_idx < len(rows) else None
                subtag = (
                    columns[start_col + cell_idx] if start_col + cell_idx < len(columns) else None
                )
                if row is None or subtag is None:
                    rejected.append(
                        {
                            "row": anchor_instance,
                            "column": subtag or "?",
                            "reason": "paste extends past the edge of the grid",
                        }
                    )
                    continue
                col_def = next(c for c in data["columns"] if c["subtag"] == subtag)
                ok, value, reason = parse_value(cell, col_def["data_type"])
                if not ok:
                    rejected.append({"row": row["name"], "column": subtag, "reason": reason})
                    continue
                old_value = self.staged_value(aoi, row["name"], subtag)
                self.upsert_staged(aoi, row["name"], subtag, old_value, value)
                applied.append({"row": row["name"], "column": subtag})
        return {"applied": applied, "rejected": rejected, "staged": len(self.staged)}

    def apply_write(self, plc_name: str) -> dict:
        """Apply staged changes to the mock PLC baseline (write = verified here)."""
        if plc_name != PLC_NAME:
            return {
                "results": [],
                "summary": {"verified": 0, "mismatch": 0, "failed": len(self.staged)},
            }
        results = []
        for entry in self.staged:
            for row in self.plc_data[entry["aoi"]]["rows"]:
                if row["name"] == entry["instance"]:
                    row["values"][entry["subtag"]] = entry["new_value"]
            results.append({"tag": entry["tag"], "status": "verified"})
        import datetime

        self.audit.append(
            {
                "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat(),
                "plc_name": PLC_NAME,
                "action": "write",
                "verified": len(results),
                "mismatch": 0,
                "failed": 0,
            }
        )
        self.staged = []
        return {
            "results": results,
            "summary": {"verified": len(results), "mismatch": 0, "failed": 0},
        }


STATE = MockState()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt: str, *args: object) -> None:
        pass

    # -- helpers ------------------------------------------------------
    def _send_json(self, status: int, payload: object) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path) -> None:
        if not path.is_file():
            self._send_json(404, {"detail": "not found"})
            return
        ctype, _ = mimetypes.guess_type(str(path))
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ctype or "application/octet-stream")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        if length == 0:
            return {}
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return {}

    # -- routing ------------------------------------------------------
    def do_GET(self) -> None:  # noqa: N802 - stdlib handler name
        path = self.path.split("?", 1)[0]
        with STATE.lock:
            if path == "/api/health":
                self._send_json(200, {"status": "ok"})
            elif path == "/api/plc":
                self._send_json(
                    200,
                    {
                        "connected": STATE.connected,
                        "plc_name": PLC_NAME if STATE.connected else None,
                        "commpath": STATE.commpath if STATE.connected else None,
                    },
                )
            elif path == "/api/aois":
                self._send_json(200, copy.deepcopy(AOIS))
            elif path.startswith("/api/aois/"):
                aoi = path[len("/api/aois/") :]
                if aoi in STATE.plc_data:
                    self._send_json(200, copy.deepcopy(STATE.plc_data[aoi]))
                else:
                    self._send_json(404, {"detail": "unknown AOI: " + aoi})
            elif path == "/api/changes":
                self._send_json(200, copy.deepcopy(STATE.staged))
            elif path == "/api/audit":
                self._send_json(200, copy.deepcopy(STATE.audit))
            elif path == "/api/export":
                self._send_json(501, {"detail": "export is coming in Phase 6"})
            elif path in ("/", "/index.html"):
                self._send_file(STATIC_DIR / "index.html")
            elif path.startswith("/vendor/") or path in ("/app.js", "/styles.css"):
                target = (STATIC_DIR / path.lstrip("/")).resolve()
                if str(target).startswith(str(STATIC_DIR.resolve())):
                    self._send_file(target)
                else:
                    self._send_json(404, {"detail": "not found"})
            else:
                self._send_json(404, {"detail": "not found"})

    def do_POST(self) -> None:  # noqa: N802 - stdlib handler name
        path = self.path.split("?", 1)[0]
        body = self._read_json()
        with STATE.lock:
            if path == "/api/connect":
                STATE.connected = True
                STATE.commpath = str(body.get("commpath") or DEFAULT_COMMPATH)
                self._send_json(
                    200,
                    {"connected": True, "plc_name": PLC_NAME, "commpath": STATE.commpath},
                )
            elif path == "/api/disconnect":
                STATE.connected = False
                self._send_json(200, {"connected": False})
            elif path.startswith("/api/aois/") and path.endswith("/paste"):
                aoi = path[len("/api/aois/") : -len("/paste")]
                if aoi not in STATE.plc_data:
                    self._send_json(404, {"detail": "unknown AOI: " + aoi})
                else:
                    result = STATE.apply_paste(
                        aoi,
                        str(body.get("tsv") or ""),
                        body.get("anchor_instance"),
                        body.get("anchor_subtag"),
                    )
                    self._send_json(200, result)
            elif path == "/api/write":
                self._send_json(200, STATE.apply_write(str(body.get("plc_name") or "")))
            else:
                self._send_json(404, {"detail": "not found"})

    def do_DELETE(self) -> None:  # noqa: N802 - stdlib handler name
        path = self.path.split("?", 1)[0]
        with STATE.lock:
            if path == "/api/changes":
                discarded = len(STATE.staged)
                STATE.staged = []
                self._send_json(200, {"discarded": discarded})
            else:
                self._send_json(404, {"detail": "not found"})


def main() -> None:
    parser = argparse.ArgumentParser(description="Dev mock API server for the PyPaxTagTool web UI.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"PyPaxTagTool mock UI at http://{args.host}:{args.port}/  (Ctrl+C to stop)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
