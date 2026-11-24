"""Tests for Phase 5 write safety: read-back verification, the discard
endpoint, and the audit file for a full write cycle."""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax.config import AppSettings
from pypax.web import create_app

PLC_NAME = "TestPLC"


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
def audit_dir(tmp_path):
    return tmp_path / "audit"


@pytest.fixture()
def app(workbook, audit_dir):
    return create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        audit_dir=str(audit_dir),
    )


@pytest.fixture()
def client(app):
    return TestClient(app)


@pytest.fixture()
def connected(client):
    response = client.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    assert response.status_code == 200
    assert response.json()["connected"] is True
    return client


def _stage_two_changes(connected):
    response = connected.post("/api/aois/P_AIn/paste", json={"tsv": "New desc\t95\t\n"})
    assert response.status_code == 200
    assert response.json()["applied"] == 2


def test_mismatch_stays_staged(connected, app):
    _stage_two_changes(connected)
    plc = app.state.pypax.plc
    plc.corrupt_after_write = {"FIT_101.Cfg_HiLim"}
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 1, "failed": 0}
    assert body["results"] == [
        {"tag": "FIT_101.Cfg_Desc", "status": "verified"},
        {"tag": "FIT_101.Cfg_HiLim", "status": "mismatch"},
    ]
    # The mismatched change stays staged for review/retry; the verified
    # one is gone and the baseline moved for it only.
    assert [c["tag"] for c in connected.get("/api/changes").json()] == ["FIT_101.Cfg_HiLim"]
    rows = connected.get("/api/aois/P_AIn").json()["rows"]
    assert rows[0]["values"][".Cfg_Desc"] == "New desc"
    assert rows[0]["values"][".Cfg_HiLim"] == 90
    # The PLC really holds the corrupted value, not the intended one.
    assert plc.read("FIT_101.Cfg_HiLim").value == 96.0
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes[0]["results"] == {"verified": 1, "mismatch": 1, "failed": 0}


def test_mismatching_tag_is_re_read_once_by_default(connected, app):
    _stage_two_changes(connected)
    plc = app.state.pypax.plc
    plc.corrupt_after_write = {"FIT_101.Cfg_HiLim"}
    connected.post("/api/write", json={"plc_name": PLC_NAME})
    # The read-back re-read the mismatching tag alone (the retry) after
    # the first read-back of both written tags.
    assert plc.read_log[-1] == ("FIT_101.Cfg_HiLim",)


def test_no_read_back_retry_when_disabled(workbook, audit_dir, tmp_path):
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        settings=AppSettings(read_back_retries=0),
        audit_dir=str(audit_dir),
    )
    client = TestClient(app)
    assert client.post("/api/connect", json={"commpath": "10.0.0.1/0"}).json()["connected"]
    _stage_two_changes(client)
    plc = app.state.pypax.plc
    plc.corrupt_after_write = {"FIT_101.Cfg_HiLim"}
    body = client.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 1, "failed": 0}
    # No single-tag re-read: the retry is disabled.
    assert ("FIT_101.Cfg_HiLim",) not in plc.read_log


def test_unreadable_read_back_is_mismatch(connected, app):
    _stage_two_changes(connected)
    plc = app.state.pypax.plc
    plc.drop_after_write = {"FIT_101.Cfg_HiLim"}
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 1, "failed": 0}
    assert body["results"] == [
        {"tag": "FIT_101.Cfg_Desc", "status": "verified"},
        {"tag": "FIT_101.Cfg_HiLim", "status": "mismatch"},
    ]
    assert [c["tag"] for c in connected.get("/api/changes").json()] == ["FIT_101.Cfg_HiLim"]


def test_already_at_value_verifies_without_writing(connected, app):
    _stage_two_changes(connected)
    plc = app.state.pypax.plc
    # Someone (or something) else already set the PLC to the staged values.
    plc._values["FIT_101.Cfg_Desc"] = "New desc"
    plc._values["FIT_101.Cfg_HiLim"] = 95.0
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 2, "mismatch": 0, "failed": 0}
    assert body["results"] == [
        {"tag": "FIT_101.Cfg_Desc", "status": "verified"},
        {"tag": "FIT_101.Cfg_HiLim", "status": "verified"},
    ]
    assert plc.write_log == []  # no write was issued at all
    assert connected.get("/api/changes").json() == []


def test_discard_changes(connected):
    _stage_two_changes(connected)
    assert len(connected.get("/api/changes").json()) == 2
    assert connected.delete("/api/changes").json() == {"discarded": 2}
    assert connected.get("/api/changes").json() == []
    # The baseline is untouched: the PLC values are still the old ones.
    rows = connected.get("/api/aois/P_AIn").json()["rows"]
    assert rows[0]["values"][".Cfg_Desc"] == "Flow Transmitter 101"
    assert rows[0]["values"][".Cfg_HiLim"] == 90
    # A later write has nothing to do and records no write entry.
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body == {"results": [], "summary": {"verified": 0, "mismatch": 0, "failed": 0}}
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes == []


def test_discard_with_nothing_staged(connected):
    assert connected.delete("/api/changes").json() == {"discarded": 0}


def test_write_cycle_is_audited_to_file(connected, audit_dir):
    _stage_two_changes(connected)
    connected.post("/api/write", json={"plc_name": PLC_NAME})
    connected.post("/api/disconnect")
    day = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    path = audit_dir / f"audit-{day}.jsonl"
    lines = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]
    assert [line["action"] for line in lines] == ["connect", "write", "disconnect"]
    write_line = lines[1]
    assert write_line["plc_name"] == PLC_NAME
    assert write_line["tags"] == 2
    assert write_line["results"] == {"verified": 2, "mismatch": 0, "failed": 0}
    # The file mirrors the in-memory entries exactly.
    entries = connected.get("/api/audit").json()
    assert lines == entries
