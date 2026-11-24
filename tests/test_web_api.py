"""API integration tests: the full contract against the FakePLC."""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax.config import PlcProfile
from pypax.web import create_app

PLC_NAME = "TestPLC"


class _FailingPLC(FakePLC):
    """A PLC that cannot be opened (connect failure path)."""

    def open(self):
        raise RuntimeError("no route to host")


@pytest.fixture()
def workbook(tmp_path):
    """Small template workbook: P_AIn sheet, FIT_101 + PIT_201 rows."""
    path = tmp_path / "template.xlsx"
    return build_workbook(
        path,
        instance_rows=[
            ("FIT_101", ["Flow Transmitter 101", 90.0, 1]),
            ("PIT_201", ["Pressure Transmitter 201", 8.0, 0]),
        ],
    )


@pytest.fixture()
def app(workbook, tmp_path):
    return create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
    )


@pytest.fixture()
def client(app):
    return TestClient(app)


@pytest.fixture()
def connected(client):
    response = client.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    assert response.status_code == 200
    assert response.json() == {
        "connected": True,
        "plc_name": PLC_NAME,
        "commpath": "10.0.0.1/0",
    }
    return client


def test_health(client):
    assert client.get("/api/health").json() == {"status": "ok"}


def test_index_serves_frontend(client):
    # The Phase 4 frontend is present in the integrated tree, so "/" serves
    # the real single-page app (the placeholder only appears when the static
    # directory is absent, e.g. a backend-only checkout).
    response = client.get("/")
    assert response.status_code == 200
    assert "PyPaxTagTool" in response.text
    assert "ag-grid" in response.text


def test_plc_status_disconnected(client):
    assert client.get("/api/plc").json() == {
        "connected": False,
        "plc_name": None,
        "commpath": None,
    }


def test_connect_and_disconnect(connected):
    assert connected.get("/api/plc").json() == {
        "connected": True,
        "plc_name": PLC_NAME,
        "commpath": "10.0.0.1/0",
    }
    assert connected.post("/api/disconnect").json() == {"connected": False}
    assert connected.get("/api/plc").json()["connected"] is False


def test_connect_failure(client, workbook):
    failing = TestClient(
        create_app(
            plc_factory=lambda commpath: _FailingPLC(commpath),
            template_path=str(workbook),
        )
    )
    response = failing.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    assert response.status_code == 200
    body = response.json()
    assert body["connected"] is False
    assert body["plc_name"] is None
    assert "no route to host" in body["error"]


def test_connect_unknown_profile(client):
    response = client.post("/api/connect", json={"commpath": "10.0.0.1/0", "profile": "nope"})
    assert response.status_code == 404
    assert "unknown profile" in response.json()["detail"]


def test_aois_require_connection(client):
    assert client.get("/api/aois").status_code == 409


def test_list_aois(connected):
    assert connected.get("/api/aois").json() == [
        {"name": "P_AIn", "instance_count": 2, "subtag_count": 3}
    ]


def test_read_aoi(connected):
    body = connected.get("/api/aois/P_AIn").json()
    assert body["name"] == "P_AIn"
    assert body["columns"] == [
        {"subtag": ".Cfg_Desc", "data_type": "STRING_40", "description": "header 5"},
        {"subtag": ".Cfg_HiLim", "data_type": "REAL", "description": "header 6"},
        {"subtag": ".HiHi.Cfg_Exists", "data_type": "BOOL", "description": "header 7"},
    ]
    assert body["rows"] == [
        {
            "name": "FIT_101",
            "values": {
                ".Cfg_Desc": "Flow Transmitter 101",
                ".Cfg_HiLim": 90,
                ".HiHi.Cfg_Exists": 1,
            },
        },
        {
            "name": "PIT_201",
            "values": {
                ".Cfg_Desc": "Pressure Transmitter 201",
                ".Cfg_HiLim": 8,
                ".HiHi.Cfg_Exists": 0,
            },
        },
    ]


def test_read_unknown_aoi(connected):
    response = connected.get("/api/aois/P_Motor")
    assert response.status_code == 404


def test_paste_positional(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste",
        json={"tsv": "New flow desc\t95\t1\nNew pressure desc\t9\t0\n"},
    )
    body = response.json()
    assert response.status_code == 200
    # 4, not 6: the pasted Exists values (1 and 0) already match the PLC,
    # so they are no-ops and are not staged.
    assert body["applied"] == 4
    assert body["rejected"] == []
    assert len(body["staged"]) == 4
    changes = connected.get("/api/changes").json()
    assert changes[0] == {
        "aoi": "P_AIn",
        "instance": "FIT_101",
        "subtag": ".Cfg_Desc",
        "tag": "FIT_101.Cfg_Desc",
        "old_value": "Flow Transmitter 101",
        "new_value": "New flow desc",
    }
    assert changes[1]["new_value"] == 95
    assert changes[1]["old_value"] == 90


def test_paste_with_instance_column(connected):
    response = connected.post("/api/aois/P_AIn/paste", json={"tsv": "PIT_201\tChanged\t9.5\t1\n"})
    body = response.json()
    assert body["applied"] == 3
    assert {c["instance"] for c in body["staged"]} == {"PIT_201"}


def test_paste_rejects_invalid_values(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste", json={"tsv": "Desc ok\tnot-a-number\tmaybe\n"}
    )
    body = response.json()
    assert body["applied"] == 1  # only the STRING survived
    assert body["rejected"] == [
        {
            "row": "FIT_101",
            "column": ".Cfg_HiLim",
            "reason": "not a valid REAL, got 'not-a-number'",
        },
        {
            "row": "FIT_101",
            "column": ".HiHi.Cfg_Exists",
            "reason": "not a valid BOOL (expected 0/1/true/false), got 'maybe'",
        },
    ]
    assert connected.get("/api/changes").json()[0]["new_value"] == "Desc ok"


def test_paste_with_anchor(connected):
    # Block starts at PIT_201 / .Cfg_HiLim: cells flow right, then down.
    response = connected.post(
        "/api/aois/P_AIn/paste",
        json={
            "tsv": "9.5\t1\n",
            "anchor_instance": "PIT_201",
            "anchor_subtag": ".Cfg_HiLim",
        },
    )
    body = response.json()
    assert body["applied"] == 2
    assert body["rejected"] == []
    staged = {(c["instance"], c["subtag"]): c["new_value"] for c in body["staged"]}
    assert staged == {("PIT_201", ".Cfg_HiLim"): 9.5, ("PIT_201", ".HiHi.Cfg_Exists"): 1}


def test_paste_with_anchor_flowing_down(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste",
        json={
            "tsv": "95\t0\n9.5\t1\n",
            "anchor_instance": "FIT_101",
            "anchor_subtag": ".Cfg_HiLim",
        },
    )
    body = response.json()
    assert body["applied"] == 4
    staged = {(c["instance"], c["subtag"]): c["new_value"] for c in body["staged"]}
    assert staged == {
        ("FIT_101", ".Cfg_HiLim"): 95,
        ("FIT_101", ".HiHi.Cfg_Exists"): 0,
        ("PIT_201", ".Cfg_HiLim"): 9.5,
        ("PIT_201", ".HiHi.Cfg_Exists"): 1,
    }


def test_paste_with_unknown_anchor_instance(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste",
        json={"tsv": "1\t2\n", "anchor_instance": "NOPE_999"},
    )
    body = response.json()
    assert body["applied"] == 0
    assert body["rejected"] == [
        {"row": "NOPE_999", "column": "", "reason": "unknown instance 'NOPE_999'"}
    ]
    assert connected.get("/api/changes").json() == []


def test_paste_with_unknown_anchor_subtag(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste",
        json={"tsv": "1\n", "anchor_subtag": ".Cfg_Nope"},
    )
    body = response.json()
    assert body["applied"] == 0
    assert body["rejected"] == [
        {"row": "FIT_101", "column": ".Cfg_Nope", "reason": "unknown subtag '.Cfg_Nope'"}
    ]


def test_paste_noop_is_not_staged(connected):
    response = connected.post(
        "/api/aois/P_AIn/paste", json={"tsv": "Flow Transmitter 101\t90\t1\n"}
    )
    assert response.json()["applied"] == 0
    assert connected.get("/api/changes").json() == []


def test_paste_skips_empty_cells(connected):
    response = connected.post("/api/aois/P_AIn/paste", json={"tsv": "\t95\t\n"})
    assert response.json()["applied"] == 1
    assert connected.get("/api/changes").json()[0]["subtag"] == ".Cfg_HiLim"


def test_paste_requires_connection(client):
    assert client.post("/api/aois/P_AIn/paste", json={"tsv": "x"}).status_code == 409


def test_write_happy_path(connected, app):
    connected.post("/api/aois/P_AIn/paste", json={"tsv": "New desc\t95\t\n"})
    response = connected.post("/api/write", json={"plc_name": PLC_NAME})
    body = response.json()
    assert response.status_code == 200
    assert body["summary"] == {"verified": 2, "mismatch": 0, "failed": 0}
    assert body["results"] == [
        {"tag": "FIT_101.Cfg_Desc", "status": "verified"},
        {"tag": "FIT_101.Cfg_HiLim", "status": "verified"},
    ]
    # Baseline updated and staged cleared; PLC really changed.
    assert connected.get("/api/changes").json() == []
    rows = connected.get("/api/aois/P_AIn").json()["rows"]
    assert rows[0]["values"][".Cfg_Desc"] == "New desc"
    assert rows[0]["values"][".Cfg_HiLim"] == 95
    audit = connected.get("/api/audit").json()
    writes = [entry for entry in audit if entry["action"] == "write"]
    assert len(writes) == 1
    assert writes[0]["plc_name"] == PLC_NAME
    assert writes[0]["tags"] == 2
    assert writes[0]["results"] == {"verified": 2, "mismatch": 0, "failed": 0}
    assert writes[0]["timestamp"]
    assert isinstance(app.state.pypax.plc, FakePLC)
    assert ("FIT_101.Cfg_Desc", "New desc") in app.state.pypax.plc.write_log


def test_write_wrong_plc_name_refused(connected):
    connected.post("/api/aois/P_AIn/paste", json={"tsv": "New desc\t95\t\n"})
    response = connected.post("/api/write", json={"plc_name": "WRONG_PLC"})
    assert response.status_code == 409
    assert "mismatch" in response.json()["detail"]
    # Nothing staged was written and the staged changes are still there.
    assert len(connected.get("/api/changes").json()) == 2
    # A refused write records no write entry (only the connect entry).
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes == []


def test_write_with_profile_expected_name_mismatch(workbook, tmp_path):
    profiles = {
        "prod": PlcProfile(
            name="prod",
            commpath="10.0.0.1/0",
            expected_plc_name="OTHER_PLC",
            safety_level="production",
        )
    }
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        profiles=profiles,
        audit_dir=str(tmp_path / "audit"),
    )
    client = TestClient(app)
    # Connect still works even though the PLC name is not the expected one.
    assert (
        client.post("/api/connect", json={"commpath": "10.0.0.1/0", "profile": "prod"}).json()[
            "connected"
        ]
        is True
    )
    client.post("/api/aois/P_AIn/paste", json={"tsv": "New desc\t95\t\n"})
    # The body names the connected PLC, but the profile expects another.
    response = client.post("/api/write", json={"plc_name": PLC_NAME})
    assert response.status_code == 409
    assert "profile" in response.json()["detail"]
    assert len(client.get("/api/changes").json()) == 2


def test_write_nothing_staged(connected):
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body == {"results": [], "summary": {"verified": 0, "mismatch": 0, "failed": 0}}
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes == []


def test_write_failed_tag_stays_staged(connected, app):
    connected.post("/api/aois/P_AIn/paste", json={"tsv": "New desc\t95\t\n"})
    # Make the PLC reject the HiLim write: drop the tag from the fake PLC
    # after staging, so the re-read sees it missing and the write fails.
    app.state.pypax.plc._values.pop("FIT_101.Cfg_HiLim")
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 0, "failed": 1}
    assert body["results"] == [
        {"tag": "FIT_101.Cfg_Desc", "status": "verified"},
        {"tag": "FIT_101.Cfg_HiLim", "status": "failed"},
    ]
    # The failed change stays staged; the written one is gone.
    assert [c["tag"] for c in connected.get("/api/changes").json()] == ["FIT_101.Cfg_HiLim"]
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes[0]["results"] == {
        "verified": 1,
        "mismatch": 0,
        "failed": 1,
    }


def test_export_not_implemented(connected):
    response = connected.get("/api/export")
    assert response.status_code == 501
    assert response.json() == {"detail": "not implemented yet"}


def test_missing_template_returns_503(tmp_path):
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(tmp_path / "nope.xlsx"),
        audit_dir=str(tmp_path / "audit"),
    )
    client = TestClient(app)
    client.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    response = client.get("/api/aois")
    assert response.status_code == 503
    assert "template workbook" in response.json()["detail"]


def test_unknown_route_is_json(client):
    response = client.get("/api/nope")
    assert response.status_code == 404
    assert "detail" in response.json()
