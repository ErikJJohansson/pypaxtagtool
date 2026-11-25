"""Static checks for the Phase 4 frontend (no browser available here).

Covers: vendored AG Grid files exist and are local (no CDN), the SPA files
exist, the JS only calls the frozen REST contract endpoints, the mock data
shapes match the contract, the write modal requires typing the PLC name,
the Export button is enabled and wired to /api/export, the Phase 8 import
UI elements and API methods exist, and the dev mock server answers the
contract end to end over real HTTP.
"""

from __future__ import annotations

import http.client
import importlib.util
import io
import json
import re
import threading
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
STATIC = ROOT / "src" / "pypax" / "web" / "static"
VENDOR = STATIC / "vendor"

# The frozen contract from the Phase 4 brief. The JS may only call these.
ALLOWED_API_PATHS = {
    "/api/health",
    "/api/connect",
    "/api/disconnect",
    "/api/plc",
    "/api/aois",
    "/api/aois/{aoi}",
    "/api/aois/{aoi}/paste",
    "/api/aois/import",
    "/api/aois/scan",
    "/api/aois/import/confirm",
    "/api/aois/custom",
    "/api/changes",
    "/api/write",
    "/api/audit",
    "/api/export",
    "/api/string-limits",
    # Phase 9a snapshots: one literal covers list/save/download/
    # compare/restore/delete (all start with /api/snapshots).
    "/api/snapshots",
}

CDN_MARKERS = ("https://", "http://", "cdn.jsdelivr.net", "unpkg.com", "cdnjs")


def _load_dev_mock_api():
    spec = importlib.util.spec_from_file_location(
        "dev_mock_api", ROOT / "scripts" / "dev_mock_api.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_frontend_files_exist():
    for path in (
        STATIC / "index.html",
        STATIC / "app.js",
        STATIC / "styles.css",
        VENDOR / "ag-grid-community.min.js",
        VENDOR / "ag-grid.css",
        VENDOR / "ag-theme-quartz.css",
        VENDOR / "README.md",
        ROOT / "scripts" / "dev_mock_api.py",
    ):
        assert path.is_file(), f"missing frontend file: {path}"


def test_ag_grid_vendored_version():
    readme = (VENDOR / "README.md").read_text(encoding="utf-8")
    assert "32.3.0" in readme
    bundle = (VENDOR / "ag-grid-community.min.js").read_text(encoding="utf-8", errors="ignore")
    assert "32.3.0" in bundle
    assert "agGrid" in bundle


def test_no_cdn_urls_in_spa():
    for name in ("index.html", "app.js", "styles.css"):
        text = (STATIC / name).read_text(encoding="utf-8")
        for marker in CDN_MARKERS:
            assert marker not in text, f"{name} references CDN marker {marker!r}"


def test_js_calls_only_frozen_endpoints():
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    literals = set(re.findall(r"'(/api/[^']*)'", js)) | set(re.findall(r'"(/api/[^"]*)"', js))
    assert literals, "no /api/ literals found in app.js — the check itself is broken"
    normalized = set()
    for raw in literals:
        path = raw.rstrip("/")
        if raw.endswith("/") and path == "/api/aois":
            normalized.add("/api/aois/{aoi}")
        else:
            normalized.add(path)
    # The paste path is built as '/api/aois/' + aoi + '/paste'.
    assert "'/paste'" in js or '"/paste"' in js
    normalized.add("/api/aois/{aoi}/paste")
    unknown = normalized - ALLOWED_API_PATHS
    assert not unknown, f"app.js calls endpoints outside the frozen contract: {unknown}"


def test_import_ui_elements_exist():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    for element_id in (
        "import-btn",
        "import-file",
        "scan-btn",
        "import-status",
        "scan-results",
        "scan-proposals",
        "scan-skipped",
        "import-modal",
        "import-meta",
        "import-nested",
        "import-columns",
        "import-select-all-btn",
        "import-select-none-btn",
        "import-cancel-btn",
        "import-confirm-btn",
        "import-modal-error",
        "remove-aoi-btn",
        "remove-modal",
        "remove-aoi-name",
        "remove-cancel-btn",
        "remove-confirm-btn",
        "remove-modal-error",
    ):
        assert f'id="{element_id}"' in html, f"missing import UI element: {element_id}"
    for token in (
        "uploadImport",
        "scanAois",
        "confirmImport",
        "removeCustomAoi",
        "MOCK_PROPOSAL",
        "custom-badge",
        "openProposalModal",
        "confirmImportProposal",
        "/api/aois/import",
        "/api/aois/scan",
        "/api/aois/import/confirm",
        "/api/aois/custom/",
    ):
        assert token in js, f"app.js missing import token: {token}"
    css = (STATIC / "styles.css").read_text(encoding="utf-8")
    assert ".custom-badge" in css
    assert ".import-table" in css


def test_snapshot_ui_elements_exist():
    """Phase 9a snapshot panel: markup, JS API methods, and CSS."""
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    for element_id in (
        "snapshots-panel",
        "snapshot-save-btn",
        "refresh-snapshots-btn",
        "snapshot-list",
        "snapshot-diff",
        "snapshot-status",
    ):
        assert f'id="{element_id}"' in html, f"missing snapshot UI element: {element_id}"
    for token in (
        "listSnapshots",
        "saveSnapshot",
        "downloadSnapshot",
        "compareSnapshot",
        "restoreSnapshot",
        "deleteSnapshot",
        "/api/snapshots",
        "renderSnapshotDiff",
        "onRestoreSnapshot",
        "auto_snapshot",
    ):
        assert token in js, f"app.js missing snapshot token: {token}"
    css = (STATIC / "styles.css").read_text(encoding="utf-8")
    assert ".snapshot-item" in css
    assert ".snapshot-diff" in css


def test_js_uses_ag_grid_create_grid():
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "agGrid.createGrid" in js
    assert "cell-changed" in js
    assert "cell-rejected" in js


def test_bool_cells_are_checkboxes_and_strings_length_checked():
    """BOOL columns render as checkboxes with AG Grid's checkbox editor;
    strings are length-checked against their type (STRING_N holds N,
    STRING holds 82, STRING_Area holds 8), mirroring the spreadsheet.
    Learned sizes arrive from GET /api/string-limits at startup
    (``learnedLimits``), with the pinned map as fallback."""
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    for token in (
        "agCheckboxCellEditor",
        "singleClickEdit",
        'type="checkbox"',
        "stringLimit",
        "-character limit for",
        "STRING_AREA",
        "STRING_NAVTAG",
        "learnedLimits",
        "getStringLimits",
        "/api/string-limits",
    ):
        assert token in js, f"app.js missing token: {token}"
    assert "includes('BOOL')" in js


def test_export_button_enabled_and_wired():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    export = re.search(r'<button[^>]*id="export-btn"[^>]*>', html)
    assert export is not None
    assert "disabled" not in export.group(0)
    assert "Export .xlsx" in html
    assert "coming in Phase 6" not in html
    assert "exportSpreadsheet" in js
    assert "fetch('/api/export')" in js
    assert "Content-Disposition" in js


def test_connect_sends_no_profile_when_field_is_empty():
    """An empty profile field must mean 'no profile' (null), not 'default'.

    The backend 404s on an unknown profile, and no 'default' profile
    exists unless the user configured one — sending 'default' made
    connecting impossible out of the box.
    """
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "|| null" in js
    assert "|| 'default'" not in js
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    assert 'placeholder="default"' not in html


def test_open_tab_reports_load_failures():
    """A tab whose PLC read fails must say so instead of hanging on
    'Loading ...' with an empty grid (the getAoi call used to reject
    unhandled)."""
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "Could not load" in js


def test_tab_highlight_matches_exact_aoi_name():
    """The active tab is matched by exact AOI name, not by prefix —
    P_AIn and P_AInAdv used to highlight together."""
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "btn.dataset.aoi" in js
    assert "btn.dataset.aoi === aoi" in js
    assert "textContent.startsWith(aoi)" not in js


def test_write_modal_requires_typing_plc_name():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert 'id="confirm-plc-input"' in html
    assert 'id="modal-confirm-btn"' in html
    # The confirm button only enables when the typed name matches state.plcName.
    assert "confirm-plc-input" in js
    assert "state.plcName" in js


def test_hidden_attribute_wins_over_display_rules():
    """Regression: a `[hidden]` element must actually be invisible.

    `.modal-overlay` sets `display: flex`, which beats the browser's own
    `[hidden] { display: none }` rule, so without a guard in our stylesheet
    every modal (write/import/remove) rendered visible on load and the
    overlay intercepted all clicks — the app was unusable until this guard.
    """
    css = (STATIC / "styles.css").read_text(encoding="utf-8")
    assert re.search(r"\[hidden\]\s*\{[^}]*display:\s*none", css), (
        "styles.css needs a [hidden] { display: none } guard; "
        "author display rules (e.g. .modal-overlay flex) override the "
        "browser's built-in hidden handling"
    )
    # The guard exists because this class sets display; if the class ever
    # stops setting display, the guard is still harmless — keep both.
    assert "display: flex" in css


def test_mock_data_shapes_match_contract():
    mock = _load_dev_mock_api()
    for aoi in mock.AOIS:
        assert set(aoi) >= {"name", "instance_count", "subtag_count"}
    for name, data in mock.PLC_DATA.items():
        assert data["name"] == name
        subtags = [c["subtag"] for c in data["columns"]]
        assert subtags, f"{name} has no columns"
        for col in data["columns"]:
            assert set(col) >= {"subtag", "data_type"}
        for row in data["rows"]:
            assert "name" in row and "values" in row
            assert set(row["values"]) == set(subtags), f"{name} row values do not match columns"
    # The in-browser mock carries the same canned data.
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "MOCK_PLC" in js
    assert "P_AIn" in js and "P_DIn" in js


def test_mock_server_contract_over_http():
    """Spin up scripts/dev_mock_api.py and exercise the whole contract."""
    mock = _load_dev_mock_api()
    server = mock.ThreadingHTTPServer(("127.0.0.1", 0), mock.Handler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=5)

        def call(method, path, body=None):
            conn.request(
                method,
                path,
                body=json.dumps(body) if body is not None else None,
                headers={"Content-Type": "application/json"},
            )
            res = conn.getresponse()
            return res.status, json.loads(res.read().decode("utf-8"))

        status, body = call("GET", "/api/health")
        assert status == 200 and body == {"status": "ok"}

        # Static files are HTML/JS, not JSON — check status + size raw.
        for static_path, min_size in (("/", 1000), ("/vendor/ag-grid-community.min.js", 100000)):
            conn.request("GET", static_path)
            res = conn.getresponse()
            payload = res.read()
            assert res.status == 200, static_path
            assert len(payload) > min_size, static_path

        status, body = call(
            "POST", "/api/connect", {"commpath": "10.10.17.10/4", "profile": "default"}
        )
        assert status == 200 and body["connected"] is True
        assert body["plc_name"] == "MOCK_PLC"

        status, aois = call("GET", "/api/aois")
        assert status == 200 and len(aois) == 2

        status, body = call("GET", "/api/aois/P_AIn")
        assert status == 200 and len(body["rows"]) == 3

        # Bad REAL value is rejected, nothing staged.
        status, body = call(
            "POST",
            "/api/aois/P_AIn/paste",
            {"tsv": "abc", "anchor_instance": "FIT-101", "anchor_subtag": ".Cfg_HiLim"},
        )
        assert status == 200
        assert body["applied"] == [] and len(body["rejected"]) == 1

        # Good value is staged.
        status, body = call(
            "POST",
            "/api/aois/P_AIn/paste",
            {"tsv": "95", "anchor_instance": "FIT-101", "anchor_subtag": ".Cfg_HiLim"},
        )
        assert status == 200 and len(body["applied"]) == 1
        # Paste returns the staged changes for the AOI (a list), like the
        # real backend — not a count.
        assert isinstance(body["staged"], list) and len(body["staged"]) == 1
        assert body["staged"][0]["tag"] == "FIT-101.Cfg_HiLim"

        status, changes = call("GET", "/api/changes")
        assert status == 200 and len(changes) == 1
        assert changes[0]["new_value"] == 95.0
        assert changes[0]["tag"] == "FIT-101.Cfg_HiLim"

        status, body = call("POST", "/api/write", {"plc_name": "MOCK_PLC"})
        assert status == 200
        assert body["summary"] == {"verified": 1, "mismatch": 0, "failed": 0}
        assert body["results"][0]["status"] == "verified"

        status, audit = call("GET", "/api/audit")
        assert status == 200 and len(audit) == 1

        # Export serves a real .xlsx (baseline + staged), not the old 501.
        conn.request("GET", "/api/export")
        res = conn.getresponse()
        payload = res.read()
        assert res.status == 200
        assert payload[:2] == b"PK"  # zip magic: a real .xlsx
        assert "MOCK_PLC_ConfigTags.xlsx" in (res.getheader("Content-Disposition") or "")
        assert "spreadsheetml.sheet" in (res.getheader("Content-Type") or "")
        import zipfile

        names = zipfile.ZipFile(io.BytesIO(payload)).namelist()
        assert "xl/workbook.xml" in names and "xl/worksheets/sheet1.xml" in names
        conn.close()
    finally:
        server.shutdown()
        server.server_close()


def test_mock_server_discard_changes():
    mock = _load_dev_mock_api()
    server = mock.ThreadingHTTPServer(("127.0.0.1", 0), mock.Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        conn = http.client.HTTPConnection("127.0.0.1", server.server_address[1], timeout=5)

        def call(method, path, body=None):
            conn.request(
                method,
                path,
                body=json.dumps(body) if body is not None else None,
                headers={"Content-Type": "application/json"},
            )
            res = conn.getresponse()
            return res.status, json.loads(res.read().decode("utf-8"))

        # Stage a change, then discard it via DELETE /api/changes.
        status, body = call(
            "POST",
            "/api/aois/P_AIn/paste",
            {"tsv": "95", "anchor_instance": "FIT-101", "anchor_subtag": ".Cfg_HiLim"},
        )
        assert status == 200 and len(body["applied"]) == 1

        status, body = call("DELETE", "/api/changes")
        assert status == 200 and body == {"discarded": 1}

        status, changes = call("GET", "/api/changes")
        assert status == 200 and changes == []

        # Discarding with nothing staged is a no-op, not an error.
        status, body = call("DELETE", "/api/changes")
        assert status == 200 and body == {"discarded": 0}
        conn.close()
    finally:
        server.shutdown()
        server.server_close()


@pytest.mark.parametrize("path", ["/api/health", "/api/aois", "/api/changes", "/api/audit"])
def test_mock_server_get_endpoints_shape(path):
    mock = _load_dev_mock_api()
    server = mock.ThreadingHTTPServer(("127.0.0.1", 0), mock.Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        conn = http.client.HTTPConnection("127.0.0.1", server.server_address[1], timeout=5)
        conn.request("GET", path)
        res = conn.getresponse()
        assert res.status == 200
        json.loads(res.read().decode("utf-8"))
        conn.close()
    finally:
        server.shutdown()
        server.server_close()


def _multipart_body(boundary: str, filename: str, content: bytes) -> bytes:
    return (
        (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            "Content-Type: application/octet-stream\r\n\r\n"
        ).encode("utf-8")
        + content
        + f"\r\n--{boundary}--\r\n".encode("utf-8")
    )


def test_mock_server_import_flow_over_http():
    """Phase 8 import endpoints: upload, scan, confirm, 409, remove."""
    mock = _load_dev_mock_api()
    server = mock.ThreadingHTTPServer(("127.0.0.1", 0), mock.Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        conn = http.client.HTTPConnection("127.0.0.1", server.server_address[1], timeout=5)

        def call(method, path, body=None, headers=None):
            hdrs = {"Content-Type": "application/json"}
            if headers:
                hdrs.update(headers)
            conn.request(method, path, body=body, headers=hdrs)
            res = conn.getresponse()
            return res.status, json.loads(res.read().decode("utf-8"))

        # Scan proposes the canned P_MyValve and skips the colliding file.
        status, scan = call("GET", "/api/aois/scan")
        assert status == 200
        assert len(scan["proposals"]) == 1
        proposal = scan["proposals"][0]
        assert proposal["aoi"] == "P_MyValve"
        assert proposal["nested_objects"] == ["FullStall"]
        assert len(proposal["columns"]) == 6
        for col in proposal["columns"]:
            assert set(col) >= {"subtag", "type", "description", "default", "selected"}
            assert col["subtag"].split(".")[-1].startswith("Cfg_")
        assert scan["skipped"] and scan["skipped"][0]["file"] == "P_AIn.L5X"
        assert scan["skipped"][0]["reason"]

        # Multipart upload returns the same canned proposal.
        boundary = "----testboundary"
        body = _multipart_body(boundary, "P_MyValve.L5X", b"<fake l5x>")
        status, uploaded = call(
            "POST",
            "/api/aois/import",
            body,
            {"Content-Type": f"multipart/form-data; boundary={boundary}"},
        )
        assert status == 200 and uploaded["aoi"] == "P_MyValve"
        assert uploaded["proposal_id"] == proposal["proposal_id"]

        # Wrong extension is a 400.
        body = _multipart_body(boundary, "notes.txt", b"nope")
        status, err = call(
            "POST",
            "/api/aois/import",
            body,
            {"Content-Type": f"multipart/form-data; boundary={boundary}"},
        )
        assert status == 400 and "detail" in err

        # Confirm with a subset of columns (drop the last one).
        selected = [c["subtag"] for c in proposal["columns"][:-1]]
        status, confirmed = call(
            "POST",
            "/api/aois/import/confirm",
            json.dumps({"proposal_id": proposal["proposal_id"], "selected": selected}),
        )
        assert status == 200
        assert confirmed == {"aoi": "P_MyValve", "columns": 5}

        # The new tab appears in the AOI list, flagged custom, and serves
        # like a template AOI with the selected columns only.
        status, aois = call("GET", "/api/aois")
        assert status == 200
        mine = next(a for a in aois if a["name"] == "P_MyValve")
        assert mine["custom"] is True and mine["subtag_count"] == 5
        assert all(a["custom"] is False for a in aois if a["name"] != "P_MyValve")
        status, data = call("GET", "/api/aois/P_MyValve")
        assert status == 200
        assert [c["subtag"] for c in data["columns"]] == selected
        assert data["rows"] == []

        # A fresh scan reports the changed export as an update, not a
        # fresh proposal; confirming it rewrites the AOI in place
        # (all-or-nothing — the selected list is ignored).
        status, scan2 = call("GET", "/api/aois/scan")
        assert status == 200
        assert scan2["proposals"] == []
        assert len(scan2["updates"]) == 1
        update = scan2["updates"][0]
        assert update["aoi"] == "P_MyValve"
        assert update["status"] == "update_available"
        assert update["update_available"] is True
        # The first confirm dropped .FullStall.Cfg_Severity, so the v2
        # export re-adds it alongside the new column.
        assert {c["subtag"] for c in update["diff"]["added"]} == {
            ".Cfg_NewCol",
            ".FullStall.Cfg_Severity",
        }
        assert [c["subtag"] for c in update["diff"]["removed"]] == [".Cfg_HasOpenLS"]
        assert [c["subtag"] for c in update["diff"]["changed"]] == [".Cfg_OpenT"]
        assert update["diff"]["unchanged"] == 3
        changed = update["diff"]["changed"][0]
        assert changed["before"]["description"] == "Time to open (sec)"
        assert changed["after"]["description"] == "Time to fully open (sec)"
        status, confirmed = call(
            "POST",
            "/api/aois/import/confirm",
            json.dumps(
                {
                    "proposal_id": update["proposal_id"],
                    "selected": selected,
                }
            ),
        )
        assert status == 200
        assert confirmed == {"aoi": "P_MyValve", "columns": 6, "updated": True}
        status, data = call("GET", "/api/aois/P_MyValve")
        assert status == 200
        assert [c["subtag"] for c in data["columns"]] == [
            ".Cfg_Desc",
            ".Cfg_Label",
            ".Cfg_OpenT",
            ".Cfg_NewCol",
            ".FullStall.Cfg_Exists",
            ".FullStall.Cfg_Severity",
        ]

        # Unknown proposal_id is a 404, like the real backend.
        status, err = call(
            "POST",
            "/api/aois/import/confirm",
            json.dumps({"proposal_id": "nope", "selected": selected}),
        )
        assert status == 404 and "unknown proposal_id" in err["detail"]

        # Remove the custom AOI; template AOIs cannot be removed.
        status, removed = call("DELETE", "/api/aois/custom/P_MyValve")
        assert status == 200 and removed == {"removed": "P_MyValve"}
        status, aois = call("GET", "/api/aois")
        assert all(a["name"] != "P_MyValve" for a in aois)
        status, err = call("DELETE", "/api/aois/custom/P_MyValve")
        assert status == 404
        status, err = call("DELETE", "/api/aois/custom/P_AIn")
        assert status == 404
        conn.close()
    finally:
        server.shutdown()
        server.server_close()
