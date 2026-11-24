"""Static checks for the Phase 4 frontend (no browser available here).

Covers: vendored AG Grid files exist and are local (no CDN), the SPA files
exist, the JS only calls the frozen REST contract endpoints, the mock data
shapes match the contract, the write modal requires typing the PLC name,
the Export button stays disabled until Phase 6, and the dev mock server
answers the contract end to end over real HTTP.
"""

from __future__ import annotations

import http.client
import importlib.util
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
    "/api/changes",
    "/api/write",
    "/api/audit",
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


def test_js_uses_ag_grid_create_grid():
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert "agGrid.createGrid" in js
    assert "cell-changed" in js
    assert "cell-rejected" in js


def test_export_button_disabled_until_phase_6():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    export = re.search(r'<button[^>]*id="export-btn"[^>]*>', html)
    assert export is not None
    assert "disabled" in export.group(0)
    assert "Phase 6" in html


def test_write_modal_requires_typing_plc_name():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    js = (STATIC / "app.js").read_text(encoding="utf-8")
    assert 'id="confirm-plc-input"' in html
    assert 'id="modal-confirm-btn"' in html
    # The confirm button only enables when the typed name matches state.plcName.
    assert "confirm-plc-input" in js
    assert "state.plcName" in js


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

        status, changes = call("GET", "/api/changes")
        assert status == 200 and len(changes) == 1
        assert changes[0]["new_value"] == 95.0
        assert changes[0]["tag"] == "FIT-101.Cfg_HiLim"

        status, body = call("POST", "/api/write", {"plc_name": "MOCK_PLC"})
        assert status == 200
        assert body["summary"] == {"written": 1, "failed": 0}
        assert body["results"][0]["status"] == "written"

        status, audit = call("GET", "/api/audit")
        assert status == 200 and len(audit) == 1

        status, _ = call("GET", "/api/export")
        assert status == 501
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
