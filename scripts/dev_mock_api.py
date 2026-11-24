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
(P_AIn with 3 instances, P_DIn with 2), PLC name ``MOCK_PLC``. The custom
AOI import endpoints (Phase 8) answer with the same canned proposal as
``MOCK_PROPOSAL`` in ``app.js``: P_MyValve with one nested-object group.
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import mimetypes
import re
import threading
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.sax.saxutils import escape

STATIC_DIR = Path(__file__).resolve().parents[1] / "src" / "pypax" / "web" / "static"

PLC_NAME = "MOCK_PLC"
DEFAULT_COMMPATH = "10.10.17.10/4"
MAX_IMPORT_BYTES = 5 * 1024 * 1024

AOIS = [
    {"name": "P_AIn", "instance_count": 3, "subtag_count": 4, "custom": False},
    {"name": "P_DIn", "instance_count": 2, "subtag_count": 3, "custom": False},
]

# Canned Phase 8 import proposal (mirrors MOCK_PROPOSAL in app.js).
MOCK_PROPOSAL = {
    "aoi": "P_MyValve",
    "revision": "1.0.0",
    "source_file": "P_MyValve.L5X",
    "columns": [
        {
            "subtag": ".Cfg_Desc",
            "type": "STRING_40",
            "description": "Description for display on HMI",
            "default": "My Valve",
            "selected": True,
        },
        {
            "subtag": ".Cfg_Label",
            "type": "STRING_20",
            "description": "Label for graphic symbol",
            "default": "My Valve",
            "selected": True,
        },
        {
            "subtag": ".Cfg_HasOpenLS",
            "type": "BOOL",
            "description": "1=Has open limit switch",
            "default": 0,
            "selected": True,
        },
        {
            "subtag": ".Cfg_OpenT",
            "type": "REAL",
            "description": "Time to open (sec)",
            "default": 5,
            "selected": True,
        },
        {
            "subtag": ".FullStall.Cfg_Exists",
            "type": "BOOL",
            "description": "1=Full Stall Alarm exists",
            "default": 0,
            "selected": True,
        },
        {
            "subtag": ".FullStall.Cfg_Severity",
            "type": "DINT",
            "description": "Alarm Severity",
            "default": 750,
            "selected": True,
        },
    ],
    "nested_objects": ["FullStall"],
    "proposal_id": "mock-proposal-1",
}

SCAN_SKIPPED = [
    {"file": "P_AIn.L5X", "reason": "AOI 'P_AIn' already exists as a template AOI"},
]

# A changed export of the canned valve, for mock-mode re-import: one
# description changed (.Cfg_OpenT), one column added (.Cfg_NewCol),
# one removed (.Cfg_HasOpenLS). Mirrors MOCK_PROPOSAL_V2 in app.js.
MOCK_PROPOSAL_V2 = {
    "aoi": "P_MyValve",
    "revision": "1.1.0",
    "source_file": "P_MyValve.L5X",
    "columns": [
        {
            "subtag": ".Cfg_Desc",
            "type": "STRING_40",
            "description": "Description for display on HMI",
            "default": "My Valve",
            "selected": True,
        },
        {
            "subtag": ".Cfg_Label",
            "type": "STRING_20",
            "description": "Label for graphic symbol",
            "default": "My Valve",
            "selected": True,
        },
        {
            "subtag": ".Cfg_OpenT",
            "type": "REAL",
            "description": "Time to fully open (sec)",
            "default": 5,
            "selected": True,
        },
        {
            "subtag": ".Cfg_NewCol",
            "type": "REAL",
            "description": "A newly added config value",
            "default": 1.5,
            "selected": True,
        },
        {
            "subtag": ".FullStall.Cfg_Exists",
            "type": "BOOL",
            "description": "1=Full Stall Alarm exists",
            "default": 0,
            "selected": True,
        },
        {
            "subtag": ".FullStall.Cfg_Severity",
            "type": "DINT",
            "description": "Alarm Severity",
            "default": 750,
            "selected": True,
        },
    ],
    "nested_objects": ["FullStall"],
    "proposal_id": "mock-proposal-2",
}


def _update_diff(before_columns: list, after_columns: list) -> dict:
    """Diff confirmed columns against a new proposal's columns.

    Same rules as the backend's diff_columns (type, description,
    default); the mock server's columns carry all three after confirm.
    """

    def as_col(column: dict, subtag: str) -> dict:
        return {
            "subtag": subtag,
            "type": column.get("type") or column.get("data_type") or "",
            "description": column.get("description") or "",
            "default": column.get("default"),
        }

    before = {column["subtag"]: column for column in before_columns}
    added, removed, changed = [], [], []
    unchanged = 0
    for column in after_columns:
        old = before.get(column["subtag"])
        if old is None:
            added.append(as_col(column, column["subtag"]))
            continue
        before_col, after_col = as_col(old, column["subtag"]), as_col(column, column["subtag"])
        if (
            before_col["type"] == after_col["type"]
            and before_col["description"] == after_col["description"]
            and before_col["default"] == after_col["default"]
        ):
            unchanged += 1
        else:
            changed.append({"subtag": column["subtag"], "before": before_col, "after": after_col})
    for subtag, column in before.items():
        if subtag not in {c["subtag"] for c in after_columns}:
            removed.append(as_col(column, subtag))
    return {"added": added, "removed": removed, "changed": changed, "unchanged": unchanged}


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
    if not dtype:
        return (
            False,
            None,
            "column type is unknown (the source .L5X is missing) — "
            "re-import the .L5X to restore type checking",
        )
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
        self._proposals: dict[str, dict] = {}
        self.custom_aois: set[str] = set()
        # Snapshots (Phase 9a): name -> {"snapshot": dict, "auto": bool}.
        self.snapshots: dict[str, dict] = {}

    def aoi_list(self) -> list[dict]:
        """GET /api/aois, derived from plc_data so imports show up."""
        return [
            {
                "name": name,
                "instance_count": len(data["rows"]),
                "subtag_count": len(data["columns"]),
                "custom": name in self.custom_aois,
            }
            for name, data in self.plc_data.items()
        ]

    def canned_proposal(self) -> dict:
        """A fresh copy of the canned proposal, registered for confirm.

        Once the AOI is imported, the changed v2 export is proposed
        instead, as an update with a column diff.
        """
        if MOCK_PROPOSAL["aoi"] in self.custom_aois:
            proposal = copy.deepcopy(MOCK_PROPOSAL_V2)
            self._proposals[proposal["proposal_id"]] = proposal
            return self.update_response(proposal)
        proposal = copy.deepcopy(MOCK_PROPOSAL)
        self._proposals[proposal["proposal_id"]] = proposal
        return proposal

    def update_response(self, proposal: dict) -> dict:
        """An update proposal: full proposal + status + column diff."""
        current = self.plc_data.get(proposal["aoi"])
        diff = _update_diff(current["columns"] if current else [], proposal["columns"])
        response = copy.deepcopy(proposal)
        response["status"] = "update_available"
        response["update_available"] = True
        response["diff"] = diff
        return response

    def scan_result(self) -> dict:
        """GET /api/aois/scan: one canned proposal, one skipped file.

        Once the AOI is imported, the scan reports the changed v2
        export under ``updates`` instead of ``proposals``.
        """
        if MOCK_PROPOSAL["aoi"] in self.custom_aois:
            proposal = copy.deepcopy(MOCK_PROPOSAL_V2)
            self._proposals[proposal["proposal_id"]] = proposal
            return {
                "proposals": [],
                "skipped": copy.deepcopy(SCAN_SKIPPED),
                "updates": [self.update_response(proposal)],
            }
        return {
            "proposals": [self.canned_proposal()],
            "skipped": copy.deepcopy(SCAN_SKIPPED),
            "updates": [],
        }

    def confirm_import(self, proposal_id: str, selected: list) -> tuple[int, dict]:
        """POST /api/aois/import/confirm. Returns (status, payload)."""
        proposal = self._proposals.get(proposal_id)
        if proposal is None:
            return 404, {"detail": "unknown proposal_id: " + str(proposal_id)}
        name = proposal["aoi"]
        if name in self.custom_aois:
            # Re-import: all-or-nothing update, selected is ignored.
            columns = proposal["columns"]
            self.plc_data[name] = {
                "name": name,
                "columns": [
                    {
                        "subtag": c["subtag"],
                        "data_type": c["type"],
                        "description": c["description"],
                        "default": c.get("default"),
                    }
                    for c in columns
                ],
                "rows": self.plc_data[name]["rows"],
            }
            del self._proposals[proposal_id]
            return 200, {"aoi": name, "columns": len(columns), "updated": True}
        if name in self.plc_data:
            return 409, {"detail": f"AOI {name!r} already exists"}
        chosen = [c for c in proposal["columns"] if c["subtag"] in set(selected or [])]
        if len(chosen) != len(selected or []):
            return 400, {"detail": "selected contains unknown subtags"}
        if not chosen:
            return 400, {"detail": "no columns selected"}
        self.plc_data[name] = {
            "name": name,
            "columns": [
                {
                    "subtag": c["subtag"],
                    "data_type": c["type"],
                    "description": c["description"],
                    "default": c.get("default"),
                }
                for c in chosen
            ],
            "rows": [],
        }
        self.custom_aois.add(name)
        del self._proposals[proposal_id]
        return 200, {"aoi": name, "columns": len(chosen)}

    def remove_custom_aoi(self, name: str) -> tuple[int, dict]:
        """DELETE /api/aois/custom/{name}. Returns (status, payload)."""
        if name not in self.custom_aois:
            return 404, {"detail": "unknown custom AOI: " + name}
        del self.plc_data[name]
        self.custom_aois.discard(name)
        return 200, {"removed": name}

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
        lines = str(tsv).splitlines()
        # Leading/trailing blank lines are not part of the block; a blank
        # line inside it is an empty row that still advances the row
        # position (matches pypax.web.paste.parse_tsv).
        while lines and not lines[0].strip():
            lines.pop(0)
        while lines and not lines[-1].strip():
            lines.pop()
        for line_idx, line in enumerate(lines):
            if not line.strip():
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
        # Same shape as the real paste endpoint: the staged changes for
        # this AOI (a list), not a count.
        staged_here = [copy.deepcopy(s) for s in self.staged if s["aoi"] == aoi]
        return {"applied": applied, "rejected": rejected, "staged": staged_here}

    def apply_write(self, plc_name: str) -> dict:
        """Apply staged changes to the mock PLC baseline (write = verified here)."""
        if plc_name != PLC_NAME:
            return {
                "results": [],
                "summary": {"verified": 0, "mismatch": 0, "failed": len(self.staged)},
            }
        # Auto-snapshot the pre-write baseline (like the real backend),
        # keeping the newest 10 auto-snapshots for this PLC.
        import datetime

        auto_name = self._take_auto_snapshot()
        results = []
        for entry in self.staged:
            for row in self.plc_data[entry["aoi"]]["rows"]:
                if row["name"] == entry["instance"]:
                    row["values"][entry["subtag"]] = entry["new_value"]
            results.append({"tag": entry["tag"], "status": "verified"})
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
            "auto_snapshot": {"ok": True, "name": auto_name},
        }

    # -- PLC snapshots (Phase 9a, in-memory stubs) --------------------

    def _snapshot_baseline(self) -> dict:
        """Current mock PLC values, in snapshot ``aois`` shape."""
        return {
            name: {row["name"]: dict(row["values"]) for row in data["rows"]}
            for name, data in self.plc_data.items()
        }

    def _new_snapshot(self) -> tuple[dict, str]:
        """A fresh snapshot dict + its auto name ``<plc>_<YYYYMMDD-HHMMSS>``."""
        import datetime

        now = datetime.datetime.now(datetime.timezone.utc)
        return {
            "format": "pypax-snapshot/1",
            "plc_name": PLC_NAME,
            "created_at": now.isoformat(),
            "aoi_version": "4.10.6",
            "aois": self._snapshot_baseline(),
        }, PLC_NAME + "_" + now.strftime("%Y%m%d-%H%M%S")

    def _store_snapshot(self, snapshot: dict, name: str, auto: bool) -> str:
        """Store a snapshot under ``name`` (suffixing same-second auto
        collisions). Returns the stored name."""
        stem = name
        suffix = 2
        while stem in self.snapshots:
            stem = f"{name}-{suffix}"
            suffix += 1
        self.snapshots[stem] = {"snapshot": snapshot, "auto": auto}
        return stem

    def _take_auto_snapshot(self) -> str:
        """Capture the pre-write baseline; prune to the newest 10."""
        snapshot, auto_name = self._new_snapshot()
        stem = self._store_snapshot(snapshot, auto_name, auto=True)
        auto_names = [
            name
            for name, entry in self.snapshots.items()
            if entry["auto"] and entry["snapshot"]["plc_name"] == PLC_NAME
        ]
        auto_names.sort(key=lambda n: self.snapshots[n]["snapshot"]["created_at"])
        for old in auto_names[:-10]:
            del self.snapshots[old]
        return stem

    def _snapshot_entry(self, name: str) -> dict:
        snapshot = self.snapshots[name]["snapshot"]
        return {
            "name": name,
            "plc_name": snapshot["plc_name"],
            "created_at": snapshot["created_at"],
            "aoi_count": len(snapshot["aois"]),
            "instance_count": sum(len(v) for v in snapshot["aois"].values()),
            "auto": self.snapshots[name]["auto"],
        }

    def list_snapshots(self) -> list[dict]:
        entries = [self._snapshot_entry(name) for name in self.snapshots]
        entries.sort(key=lambda e: (e["created_at"], e["name"]), reverse=True)
        return entries

    def save_snapshot(self, name: str | None) -> tuple[int, dict]:
        """POST /api/snapshots. Returns (status, payload)."""
        if name is not None and not re.fullmatch(r"[A-Za-z0-9_-]+", name):
            return 400, {"detail": f"snapshot name must be [A-Za-z0-9_-]+, got {name!r}"}
        if name is not None and name in self.snapshots:
            return 409, {"detail": f"snapshot {name!r} already exists"}
        snapshot, auto_name = self._new_snapshot()
        stem = self._store_snapshot(snapshot, name or auto_name, auto=name is None)
        return 201, self._snapshot_entry(stem)

    def _diff_snapshot(self, snapshot: dict) -> list[dict]:
        """Differences between a snapshot and the mock PLC baseline."""
        diffs = []
        for aoi in sorted(set(snapshot["aois"]) | set(self.plc_data)):
            snap_rows = snapshot["aois"].get(aoi, {})
            cur_rows = {
                row["name"]: row["values"] for row in self.plc_data.get(aoi, {}).get("rows", [])
            }
            for instance in sorted(set(snap_rows) | set(cur_rows)):
                snap_values = snap_rows.get(instance, {})
                cur_values = cur_rows.get(instance, {})
                for subtag in sorted(set(snap_values) | set(cur_values)):
                    if snap_values.get(subtag) != cur_values.get(subtag):
                        diffs.append(
                            {
                                "aoi": aoi,
                                "instance": instance,
                                "subtag": subtag,
                                "snapshot_value": snap_values.get(subtag),
                                "current_value": cur_values.get(subtag),
                            }
                        )
        return diffs

    def compare_snapshot(self, name: str) -> tuple[int, dict]:
        """POST /api/snapshots/{name}/compare. Returns (status, payload)."""
        if name not in self.snapshots:
            return 404, {"detail": f"unknown snapshot {name!r}"}
        snapshot = self.snapshots[name]["snapshot"]
        return 200, {
            "plc_name": snapshot["plc_name"],
            "differences": self._diff_snapshot(snapshot),
        }

    def restore_snapshot(self, name: str) -> tuple[int, dict]:
        """POST /api/snapshots/{name}/restore: stage differing values."""
        if name not in self.snapshots:
            return 404, {"detail": f"unknown snapshot {name!r}"}
        snapshot = self.snapshots[name]["snapshot"]
        if snapshot["plc_name"] != PLC_NAME:
            return 409, {
                "detail": f"snapshot is for PLC {snapshot['plc_name']!r}, connected to {PLC_NAME!r}"
            }
        staged = 0
        skipped = []
        for aoi, instances in snapshot["aois"].items():
            data = self.plc_data.get(aoi)
            for instance, values in instances.items():
                for subtag, snap_value in values.items():
                    entry = {"aoi": aoi, "instance": instance, "subtag": subtag}
                    if snap_value is None:
                        skipped.append({**entry, "reason": "no value in snapshot"})
                        continue
                    if data is None:
                        skipped.append({**entry, "reason": f"unknown AOI {aoi!r}"})
                        continue
                    column = next((c for c in data["columns"] if c["subtag"] == subtag), None)
                    if column is None:
                        skipped.append({**entry, "reason": f"unknown subtag {subtag!r} in {aoi}"})
                        continue
                    row = next((r for r in data["rows"] if r["name"] == instance), None)
                    if row is None:
                        skipped.append(
                            {**entry, "reason": f"instance {instance!r} not found in PLC"}
                        )
                        continue
                    ok, value, reason = parse_value(str(snap_value), column["data_type"])
                    if not ok:
                        skipped.append({**entry, "reason": reason})
                        continue
                    if value == row["values"][subtag]:
                        continue
                    self.upsert_staged(aoi, instance, subtag, row["values"][subtag], value)
                    staged += 1
        return 200, {"staged": staged, "skipped": skipped}

    def delete_snapshot(self, name: str) -> tuple[int, dict]:
        """DELETE /api/snapshots/{name}. Returns (status, payload)."""
        if name not in self.snapshots:
            return 404, {"detail": f"unknown snapshot {name!r}"}
        del self.snapshots[name]
        return 200, {"removed": name}

    def upload_snapshot(self, filename: str, data: bytes) -> tuple[int, dict]:
        """POST /api/snapshots/upload. Returns (status, payload)."""
        if not (filename or "").lower().endswith(".json"):
            return 400, {"detail": "only .json snapshot files can be uploaded"}
        if len(data) > MAX_IMPORT_BYTES:
            return 400, {"detail": "snapshot exceeds the 5 MB cap"}
        try:
            snapshot = json.loads(data.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return 400, {"detail": "snapshot is not valid JSON"}
        if not isinstance(snapshot, dict) or snapshot.get("format") != "pypax-snapshot/1":
            return 400, {"detail": "not a pypax-snapshot/1 snapshot"}
        if snapshot.get("plc_name") != PLC_NAME:
            return 409, {
                "detail": f"snapshot is for PLC {snapshot.get('plc_name')!r}, "
                f"connected to {PLC_NAME!r}"
            }
        stem = Path(filename).stem
        if not re.fullmatch(r"[A-Za-z0-9_-]+", stem):
            return 400, {"detail": f"snapshot name must be [A-Za-z0-9_-]+, got {stem!r}"}
        if stem in self.snapshots:
            return 409, {"detail": f"snapshot {stem!r} already exists"}
        self.snapshots[stem] = {"snapshot": snapshot, "auto": False}
        return 201, self._snapshot_entry(stem)


def _xlsx_col_letter(index: int) -> str:
    """1-based column index -> Excel column letters (1 -> A, 27 -> AA)."""
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _xlsx_row(row_num: int, values: list) -> str:
    """One <row> of inline-string / numeric cells."""
    cells = []
    for col_idx, value in enumerate(values, start=1):
        ref = _xlsx_col_letter(col_idx) + str(row_num)
        if isinstance(value, bool):
            cells.append(f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>')
        elif isinstance(value, (int, float)):
            cells.append(f'<c r="{ref}"><v>{value}</v></c>')
        else:
            text = escape("" if value is None else str(value))
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
    return f'<row r="{row_num}">{"".join(cells)}</row>'


def _build_mock_export(state: MockState) -> bytes:
    """A minimal valid .xlsx of the mock PLC baseline + staged changes.

    Stdlib zipfile only (the mock server has no openpyxl). One sheet
    per AOI: header row (Tag Name + subtags), one row per instance,
    staged values overlaid on the baseline. No amber fill — the mock
    is a UI stand-in, not the real export.
    """
    sheet_names = list(state.plc_data)
    sheet_xmls: dict[str, str] = {}
    for index, aoi in enumerate(sheet_names, start=1):
        data = state.plc_data[aoi]
        rows_xml = [_xlsx_row(1, ["Tag Name"] + [c["subtag"] for c in data["columns"]])]
        for row_idx, row in enumerate(data["rows"], start=2):
            values: list = [row["name"]]
            for col in data["columns"]:
                staged = next(
                    (
                        s["new_value"]
                        for s in state.staged
                        if s["aoi"] == aoi
                        and s["instance"] == row["name"]
                        and s["subtag"] == col["subtag"]
                    ),
                    None,
                )
                values.append(row["values"][col["subtag"]] if staged is None else staged)
            rows_xml.append(_xlsx_row(row_idx, values))
        sheet_xmls[f"sheet{index}"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheetData>" + "".join(rows_xml) + "</sheetData></worksheet>"
        )
    sheets_el = "".join(
        f'<sheet name="{escape(aoi)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, aoi in enumerate(sheet_names, start=1)
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_el}</sheets></workbook>"
    )
    rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ws_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{rels_ns}">'
        + "".join(
            f'<Relationship Id="rId{i}" Type="{ws_rel}" Target="worksheets/sheet{i}.xml"/>'
            for i in range(1, len(sheet_names) + 1)
        )
        + "</Relationships>"
    )
    main_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    sheet_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{main_type}"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{sheet_type}"/>'
            for i in range(1, len(sheet_names) + 1)
        )
        + "</Types>"
    )
    doc_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    package_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{rels_ns}">'
        f'<Relationship Id="rId1" Type="{doc_rel}" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", package_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for name, xml in sheet_xmls.items():
            zf.writestr(f"xl/worksheets/{name}.xml", xml)
    return buf.getvalue()


def _json_from_bytes(raw: bytes) -> dict:
    if not raw:
        return {}
    try:
        return json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return {}


def _parse_multipart(body: bytes, content_type: str) -> tuple[str | None, bytes | None, str | None]:
    """Extract the ``file`` part of a multipart/form-data body.

    Returns (filename, file_bytes, error); error is None on success. The
    file content is not inspected — the mock always answers with the
    canned proposal — but the filename and size are validated like the
    real backend (400 on bad input).
    """
    match = re.search(r"boundary=([^;]+)", content_type or "")
    if not match:
        return None, None, "not a multipart upload"
    boundary = match.group(1).strip().strip('"').encode("ascii")
    for part in body.split(b"--" + boundary):
        if b'name="file"' not in part:
            continue
        head, sep, content = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        name_match = re.search(rb'filename="([^"]*)"', head)
        filename = name_match.group(1).decode("utf-8", "replace") if name_match else ""
        if content.endswith(b"\r\n"):
            content = content[:-2]
        return filename, content, None
    return None, None, "no file uploaded"


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

    def _read_body(self) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(length) if length > 0 else b""

    def _read_json(self) -> dict:
        return _json_from_bytes(self._read_body())

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
                self._send_json(200, STATE.aoi_list())
            elif path == "/api/aois/scan":
                self._send_json(200, STATE.scan_result())
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
            elif path == "/api/snapshots":
                self._send_json(200, STATE.list_snapshots())
            elif path.startswith("/api/snapshots/"):
                name = path[len("/api/snapshots/") :]
                entry = STATE.snapshots.get(name)
                if entry is None or "/" in name:
                    self._send_json(404, {"detail": "unknown snapshot: " + name})
                else:
                    payload = json.dumps(entry["snapshot"], indent=2).encode("utf-8")
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Content-Length", str(len(payload)))
                    self.send_header("Content-Disposition", f'attachment; filename="{name}.json"')
                    self.end_headers()
                    self.wfile.write(payload)
            elif path == "/api/export":
                payload = _build_mock_export(STATE)
                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                self.send_header("Content-Length", str(len(payload)))
                self.send_header(
                    "Content-Disposition",
                    'attachment; filename="MOCK_PLC_ConfigTags.xlsx"',
                )
                self.end_headers()
                self.wfile.write(payload)
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
        raw = self._read_body()
        body = _json_from_bytes(raw)
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
            elif path == "/api/aois/import":
                filename, content, error = _parse_multipart(
                    raw, self.headers.get("Content-Type") or ""
                )
                if error is not None:
                    self._send_json(400, {"detail": error})
                elif not (filename or "").lower().endswith(".l5x"):
                    self._send_json(400, {"detail": "only .L5X files can be imported"})
                elif content is not None and len(content) > MAX_IMPORT_BYTES:
                    self._send_json(400, {"detail": "file is larger than the 5 MB import limit"})
                else:
                    self._send_json(200, STATE.canned_proposal())
            elif path == "/api/aois/import/confirm":
                status, payload = STATE.confirm_import(
                    str(body.get("proposal_id") or ""), body.get("selected") or []
                )
                self._send_json(status, payload)
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
            elif path == "/api/snapshots":
                status, payload = STATE.save_snapshot(body.get("name"))
                self._send_json(status, payload)
            elif path == "/api/snapshots/upload":
                filename, content, error = _parse_multipart(
                    raw, self.headers.get("Content-Type") or ""
                )
                if error is not None:
                    self._send_json(400, {"detail": error})
                elif content is None or len(content) > MAX_IMPORT_BYTES:
                    self._send_json(400, {"detail": "file is larger than the 5 MB import limit"})
                else:
                    status, payload = STATE.upload_snapshot(filename, content)
                    self._send_json(status, payload)
            elif path.startswith("/api/snapshots/") and path.endswith("/compare"):
                name = path[len("/api/snapshots/") : -len("/compare")]
                status, payload = STATE.compare_snapshot(name)
                self._send_json(status, payload)
            elif path.startswith("/api/snapshots/") and path.endswith("/restore"):
                name = path[len("/api/snapshots/") : -len("/restore")]
                status, payload = STATE.restore_snapshot(name)
                self._send_json(status, payload)
            else:
                self._send_json(404, {"detail": "not found"})

    def do_DELETE(self) -> None:  # noqa: N802 - stdlib handler name
        path = self.path.split("?", 1)[0]
        with STATE.lock:
            if path == "/api/changes":
                discarded = len(STATE.staged)
                STATE.staged = []
                self._send_json(200, {"discarded": discarded})
            elif path.startswith("/api/aois/custom/"):
                name = path[len("/api/aois/custom/") :]
                status, payload = STATE.remove_custom_aoi(name)
                self._send_json(status, payload)
            elif path.startswith("/api/snapshots/"):
                name = path[len("/api/snapshots/") :]
                status, payload = STATE.delete_snapshot(name)
                self._send_json(status, payload)
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
