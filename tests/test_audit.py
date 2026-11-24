"""Tests for the audit log: in-memory entries + daily JSONL files."""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax.web import create_app
from pypax.web.audit import AuditLog

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
        snapshot_dir=str(audit_dir.parent / "snapshots"),
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


def _read_lines(audit_dir):
    day = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    path = audit_dir / f"audit-{day}.jsonl"
    assert path.exists()
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]


def test_audit_log_records_to_memory_and_file(tmp_path):
    log = AuditLog(tmp_path / "audit")
    log.record(
        {
            "timestamp": "2026-11-24T12:00:00+00:00",
            "plc_name": "TestPLC",
            "action": "write",
            "tags": 2,
            "results": {"verified": 2, "mismatch": 0, "failed": 0},
        }
    )
    log.record(
        {
            "timestamp": "2026-11-24T12:01:00+00:00",
            "plc_name": "TestPLC",
            "action": "disconnect",
            "tags": 0,
            "results": {"connected": False},
        }
    )
    assert len(log.list()) == 2
    assert log.list()[0]["action"] == "write"
    lines = _read_lines(tmp_path / "audit")
    assert len(lines) == 2
    assert lines[0]["action"] == "write"
    assert lines[0]["results"] == {"verified": 2, "mismatch": 0, "failed": 0}
    assert lines[1]["action"] == "disconnect"
    # list() returns a copy: mutating it must not touch the log.
    log.list().clear()
    assert len(log.list()) == 2


def test_audit_log_survives_unwritable_directory(tmp_path):
    blocker = tmp_path / "blocker"
    blocker.write_text("not a directory", encoding="utf-8")
    log = AuditLog(blocker / "audit")
    entry = {
        "timestamp": "2026-11-24T12:00:00+00:00",
        "plc_name": "TestPLC",
        "action": "write",
        "tags": 1,
        "results": {"verified": 1, "mismatch": 0, "failed": 0},
    }
    log.record(entry)  # must not raise: the file is best-effort
    assert log.list() == [entry]


def test_connect_and_disconnect_are_audited_to_file(connected, audit_dir):
    connected.post("/api/disconnect")
    entries = connected.get("/api/audit").json()
    assert [e["action"] for e in entries] == ["connect", "disconnect"]
    assert entries[0]["plc_name"] == PLC_NAME
    assert entries[0]["tags"] == 0
    assert entries[0]["results"] == {"connected": True}
    assert entries[1]["results"] == {"connected": False}
    lines = _read_lines(audit_dir)
    assert [line["action"] for line in lines] == ["connect", "disconnect"]
    assert lines[0] == entries[0]


def test_audit_survives_disconnect(connected, audit_dir):
    connected.post("/api/disconnect")
    assert len(connected.get("/api/audit").json()) == 2  # connect + disconnect
    assert len(_read_lines(audit_dir)) == 2
