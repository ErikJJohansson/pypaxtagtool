"""Tests for Phase 9a PLC snapshots: the snapshots module (unit) and
the snapshot API (TestClient + FakePLC end to end)."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax import snapshots
from pypax.web import create_app

PLC_NAME = "TestPLC"
T0 = datetime(2026, 11, 24, 10, 0, 0, tzinfo=timezone.utc)
T1 = datetime(2026, 11, 24, 11, 0, 0, tzinfo=timezone.utc)
T2 = datetime(2026, 11, 24, 12, 0, 0, tzinfo=timezone.utc)

AOIS = {
    "P_AIn": {
        "FIT_101": {".Cfg_Desc": "Flow", ".Cfg_HiLim": 90.0, ".HiHi.Cfg_Exists": 1},
        "PIT_201": {".Cfg_Desc": "Pressure", ".Cfg_HiLim": 8.0, ".HiHi.Cfg_Exists": 0},
    }
}


# ----------------------------------------------------------------------
# Module unit tests
# ----------------------------------------------------------------------


def test_build_snapshot_shape():
    snap = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    assert snap["format"] == "pypax-snapshot/1"
    assert snap["plc_name"] == PLC_NAME
    assert snap["aoi_version"] == "4.10.6"
    assert snap["created_at"] == T0.isoformat()
    assert "+" in snap["created_at"]  # ISO 8601 with offset
    assert snap["aois"] == AOIS
    assert snapshots.snapshot_counts(snap) == (1, 2)


def test_build_snapshot_normalizes_bools_to_int():
    snap = snapshots.build_snapshot(
        PLC_NAME,
        {"P_AIn": {"FIT_101": {".HiHi.Cfg_Exists": True, ".Cfg_HiLim": 90.0}}},
        "4.10.6",
        created_at=T0,
    )
    value = snap["aois"]["P_AIn"]["FIT_101"][".HiHi.Cfg_Exists"]
    assert value == 1 and type(value) is int


def test_build_snapshot_rejects_bad_inputs():
    with pytest.raises(ValueError):
        snapshots.build_snapshot("", AOIS, "4.10.6", created_at=T0)
    with pytest.raises(ValueError):
        snapshots.build_snapshot(PLC_NAME, AOIS, "", created_at=T0)
    with pytest.raises(ValueError):
        snapshots.build_snapshot(PLC_NAME, ["not", "a", "dict"], "4.10.6", created_at=T0)
    # A naive created_at (no offset) is rejected: the format requires one.
    with pytest.raises(ValueError):
        snapshots.build_snapshot(
            PLC_NAME, AOIS, "4.10.6", created_at=datetime(2026, 11, 24, 10, 0, 0)
        )
    with pytest.raises(ValueError):
        snapshots.build_snapshot(
            PLC_NAME,
            {"P_AIn": {"FIT_101": {".Cfg_Desc": ["a", "list"]}}},
            "4.10.6",
            created_at=T0,
        )


def test_sanitize_snapshot_name():
    for good in ("before-outage", "baseline_1", "ABC123", "a"):
        assert snapshots.sanitize_snapshot_name(good) == good
    for bad in ("", "a/b", "../x", "a b", "a.b", "a:b", None, 42):
        with pytest.raises(ValueError):
            snapshots.sanitize_snapshot_name(bad)


def test_snapshot_filename_sanitizes_plc_name():
    assert snapshots.snapshot_filename("Test PLC/1", T0) == "Test_PLC_1_20261124-100000"
    assert snapshots.snapshot_filename(PLC_NAME, T0) == "TestPLC_20261124-100000"


def test_save_load_round_trip(tmp_path):
    snap = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    path = snapshots.save_snapshot(tmp_path, snap, name="before-outage")
    assert path.name == "before-outage.json"
    assert snapshots.load_snapshot(path) == snap


def test_save_auto_same_second_gets_suffix(tmp_path):
    snap = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    first = snapshots.save_snapshot(tmp_path, snap)
    second = snapshots.save_snapshot(tmp_path, snap)
    assert first.name == "TestPLC_20261124-100000.json"
    assert second.name == "TestPLC_20261124-100000-2.json"
    assert first.parent.name == "auto" and second.parent.name == "auto"


def test_save_manual_duplicate_is_refused(tmp_path):
    snap = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    snapshots.save_snapshot(tmp_path, snap, name="baseline")
    with pytest.raises(ValueError, match="already exists"):
        snapshots.save_snapshot(tmp_path, snap, name="baseline")


def test_save_manual_name_colliding_with_auto_is_refused(tmp_path):
    """F7: a manual name must not shadow an auto-snapshot of the same name.

    Lookup by name is single-valued (manual wins), so two snapshots
    sharing a stem would leave one unreachable. A manual save whose
    stem already exists as an auto-snapshot is refused.
    """
    auto = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    snapshots.save_snapshot(tmp_path, auto)
    manual = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T1)
    with pytest.raises(ValueError, match="already exists"):
        snapshots.save_snapshot(tmp_path, manual, name="TestPLC_20261124-100000")


def test_save_auto_name_colliding_with_manual_gets_suffix(tmp_path):
    """F7: an auto-snapshot whose stem a manual snapshot already has
    gets a ``-2`` suffix instead of shadowing it."""
    manual = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    snapshots.save_snapshot(tmp_path, manual, name="TestPLC_20261124-100000")
    auto = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    saved = snapshots.save_snapshot(tmp_path, auto)
    assert saved.name == "TestPLC_20261124-100000-2.json"
    assert saved.parent.name == "auto"
    assert snapshots.load_snapshot(saved)["plc_name"] == PLC_NAME


def test_list_dedupes_hand_placed_name_collision(tmp_path):
    """F7: a name shared by a manual and an auto file (only possible
    for hand-placed files — save_snapshot refuses the collision)
    lists once, the manual one, matching name lookup."""
    auto = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)
    manual = snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T1)
    (tmp_path / "auto").mkdir()
    (tmp_path / "auto" / "TestPLC_20261124-100000.json").write_text(
        json.dumps(auto), encoding="utf-8"
    )
    (tmp_path / "TestPLC_20261124-100000.json").write_text(json.dumps(manual), encoding="utf-8")
    entries = snapshots.list_snapshots(tmp_path)
    assert [entry["name"] for entry in entries] == ["TestPLC_20261124-100000"]
    assert entries[0]["auto"] is False
    assert entries[0]["created_at"] == T1.isoformat()


def test_load_rejects_corrupt_and_wrong_shape(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json", encoding="utf-8")
    with pytest.raises(ValueError, match="not valid JSON"):
        snapshots.load_snapshot(bad)
    wrong = tmp_path / "wrong.json"
    wrong.write_text(json.dumps({"format": "other/1"}), encoding="utf-8")
    with pytest.raises(ValueError, match="not a pypax-snapshot/1"):
        snapshots.load_snapshot(wrong)
    shape = tmp_path / "shape.json"
    shape.write_text(
        json.dumps(
            {
                "format": "pypax-snapshot/1",
                "plc_name": PLC_NAME,
                "created_at": T0.isoformat(),
                "aoi_version": "4.10.6",
                "aois": {"P_AIn": {"FIT_101": {".Cfg_Desc": {"nested": True}}}},
            }
        ),
        encoding="utf-8",
    )
    with pytest.raises(ValueError):
        snapshots.load_snapshot(shape)
    naive = tmp_path / "naive.json"
    naive.write_text(
        json.dumps(
            {
                "format": "pypax-snapshot/1",
                "plc_name": PLC_NAME,
                "created_at": "2026-11-24T10:00:00",
                "aoi_version": "4.10.6",
                "aois": {},
            }
        ),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="timezone offset"):
        snapshots.load_snapshot(naive)
    with pytest.raises(FileNotFoundError):
        snapshots.load_snapshot(tmp_path / "missing.json")


def test_parse_snapshot_bytes_rejects_oversize():
    data = snapshots.parse_snapshot_bytes(
        json.dumps(snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T0)).encode()
    )
    assert data["plc_name"] == PLC_NAME
    with pytest.raises(ValueError, match="not valid JSON"):
        snapshots.parse_snapshot_bytes(b"{oops")
    with pytest.raises(ValueError, match="MB cap"):
        snapshots.parse_snapshot_bytes(b" " * (snapshots.MAX_UPLOAD_BYTES + 1))


def test_list_snapshots_newest_first_and_skips_corrupt(tmp_path):
    for stamp, name in ((T0, "old"), (T2, "new"), (T1, "mid")):
        snapshots.save_snapshot(
            tmp_path,
            snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=stamp),
            name=name,
        )
    snapshots.save_snapshot(
        tmp_path, snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=T1)
    )
    (tmp_path / "corrupt.json").write_text("{broken", encoding="utf-8")
    entries = snapshots.list_snapshots(tmp_path)
    assert [entry["name"] for entry in entries] == ["new", "mid", "TestPLC_20261124-110000", "old"]
    assert [entry["auto"] for entry in entries] == [False, False, True, False]
    assert entries[0]["plc_name"] == PLC_NAME
    assert entries[0]["aoi_count"] == 1 and entries[0]["instance_count"] == 2
    assert all(entry["created_at"] for entry in entries)


def test_diff_snapshot():
    current = {
        "P_AIn": {
            "FIT_101": {".Cfg_Desc": "Flow!", ".Cfg_HiLim": 90.0, ".HiHi.Cfg_Exists": 1},
            "TIT_301": {".Cfg_Desc": "New instance", ".Cfg_HiLim": 5.0},
        }
    }
    diffs = snapshots.diff_snapshot(AOIS, current)
    assert diffs == [
        {
            "aoi": "P_AIn",
            "instance": "FIT_101",
            "subtag": ".Cfg_Desc",
            "snapshot_value": "Flow",
            "current_value": "Flow!",
        },
        {
            "aoi": "P_AIn",
            "instance": "PIT_201",
            "subtag": ".Cfg_Desc",
            "snapshot_value": "Pressure",
            "current_value": None,
        },
        {
            "aoi": "P_AIn",
            "instance": "PIT_201",
            "subtag": ".Cfg_HiLim",
            "snapshot_value": 8.0,
            "current_value": None,
        },
        {
            "aoi": "P_AIn",
            "instance": "PIT_201",
            "subtag": ".HiHi.Cfg_Exists",
            "snapshot_value": 0,
            "current_value": None,
        },
        {
            "aoi": "P_AIn",
            "instance": "TIT_301",
            "subtag": ".Cfg_Desc",
            "snapshot_value": None,
            "current_value": "New instance",
        },
        {
            "aoi": "P_AIn",
            "instance": "TIT_301",
            "subtag": ".Cfg_HiLim",
            "snapshot_value": None,
            "current_value": 5.0,
        },
    ]
    # Equal values are not differences, and identical inputs give none.
    assert snapshots.diff_snapshot(AOIS, AOIS) == []


def test_prune_auto_snapshots(tmp_path):
    for i in range(12):
        stamp = datetime(2026, 11, 24, 10, i, 0, tzinfo=timezone.utc)
        snapshots.save_snapshot(
            tmp_path, snapshots.build_snapshot(PLC_NAME, AOIS, "4.10.6", created_at=stamp)
        )
    for i in range(2):
        stamp = datetime(2026, 11, 24, 10, i, 0, tzinfo=timezone.utc)
        snapshots.save_snapshot(
            tmp_path, snapshots.build_snapshot("OtherPLC", AOIS, "4.10.6", created_at=stamp)
        )
    (tmp_path / "auto" / "corrupt.json").write_text("{broken", encoding="utf-8")
    removed = snapshots.prune_auto_snapshots(tmp_path, PLC_NAME)
    assert removed == 2
    remaining = snapshots.list_snapshots(tmp_path)
    auto_names = [entry["name"] for entry in remaining if entry["auto"]]
    assert len(auto_names) == 12  # 10 TestPLC + 2 OtherPLC; corrupt file skipped, kept
    assert "TestPLC_20261124-100000" not in auto_names
    assert "TestPLC_20261124-100100" not in auto_names
    assert "TestPLC_20261124-100200" in auto_names
    assert (tmp_path / "auto" / "corrupt.json").exists()


# ----------------------------------------------------------------------
# API tests (TestClient + FakePLC)
# ----------------------------------------------------------------------


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
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
        snapshot_dir=str(tmp_path / "snapshots"),
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


def _move_fake_values(app, desc="Changed desc", hilim=95.0, exists=False):
    plc = app.state.pypax.plc
    plc._values["FIT_101.Cfg_Desc"] = desc
    plc._values["FIT_101.Cfg_HiLim"] = hilim
    plc._values["FIT_101.HiHi.Cfg_Exists"] = exists


def test_snapshot_round_trip_compare_restore_write(connected, app):
    # Save the baseline (covers every AOI, not just opened tabs).
    response = connected.post("/api/snapshots", json={"name": "before-outage"})
    assert response.status_code == 201, response.json()
    body = response.json()
    assert body["name"] == "before-outage"
    assert body["plc_name"] == PLC_NAME
    assert body["aoi_count"] == 1 and body["instance_count"] == 2

    # The PLC moves on; compare shows exactly what changed.
    _move_fake_values(app)
    response = connected.post("/api/snapshots/before-outage/compare")
    assert response.status_code == 200
    diffs = response.json()["differences"]
    assert diffs == [
        {
            "aoi": "P_AIn",
            "instance": "FIT_101",
            "subtag": ".Cfg_Desc",
            "snapshot_value": "Flow Transmitter 101",
            "current_value": "Changed desc",
        },
        {
            "aoi": "P_AIn",
            "instance": "FIT_101",
            "subtag": ".Cfg_HiLim",
            "snapshot_value": 90,
            "current_value": 95,
        },
        {
            "aoi": "P_AIn",
            "instance": "FIT_101",
            "subtag": ".HiHi.Cfg_Exists",
            "snapshot_value": 1,
            "current_value": 0,
        },
    ]

    # Restore stages only — the PLC still holds the moved values.
    response = connected.post("/api/snapshots/before-outage/restore")
    assert response.status_code == 200
    assert response.json()["staged"] == 3
    assert response.json()["skipped"] == []
    plc = app.state.pypax.plc
    assert plc._values["FIT_101.Cfg_HiLim"] == 95.0
    staged = connected.get("/api/changes").json()
    assert [c["tag"] for c in staged] == [
        "FIT_101.Cfg_Desc",
        "FIT_101.Cfg_HiLim",
        "FIT_101.HiHi.Cfg_Exists",
    ]
    assert staged[0]["old_value"] == "Changed desc"
    assert staged[0]["new_value"] == "Flow Transmitter 101"

    # Write it back: read-back verified, PLC back at the snapshot.
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 3, "mismatch": 0, "failed": 0}
    assert plc._values["FIT_101.Cfg_HiLim"] == 90.0
    assert plc._values["FIT_101.Cfg_Desc"] == "Flow Transmitter 101"
    assert plc._values["FIT_101.HiHi.Cfg_Exists"] == 1
    assert connected.get("/api/changes").json() == []


def test_snapshot_list_download_and_delete(connected, tmp_path):
    connected.post("/api/snapshots", json={"name": "baseline"})
    entries = connected.get("/api/snapshots").json()
    assert len(entries) == 1
    assert entries[0]["name"] == "baseline"
    assert entries[0]["auto"] is False
    assert entries[0]["plc_name"] == PLC_NAME

    response = connected.get("/api/snapshots/baseline")
    assert response.status_code == 200
    assert "attachment" in response.headers["Content-Disposition"]
    downloaded = response.json()
    assert downloaded["format"] == "pypax-snapshot/1"
    assert downloaded["aois"]["P_AIn"]["FIT_101"][".Cfg_HiLim"] == 90

    response = connected.delete("/api/snapshots/baseline")
    assert response.status_code == 200
    assert response.json() == {"removed": "baseline"}
    assert connected.get("/api/snapshots").json() == []
    assert connected.get("/api/snapshots/baseline").status_code == 404


def test_snapshot_requires_connection(client):
    assert client.post("/api/snapshots", json={}).status_code == 409
    assert client.post("/api/snapshots/whatever/compare").status_code == 409
    assert client.post("/api/snapshots/whatever/restore").status_code == 409
    # The listing is a file listing: it works without a connection.
    assert client.get("/api/snapshots").status_code == 200


def test_snapshot_unknown_name_404(connected):
    assert connected.post("/api/snapshots/nope/compare").status_code == 404
    assert connected.post("/api/snapshots/nope/restore").status_code == 404
    assert connected.get("/api/snapshots/nope").status_code == 404
    assert connected.delete("/api/snapshots/nope").status_code == 404


def test_snapshot_name_validation(connected):
    assert connected.post("/api/snapshots", json={"name": "bad name"}).status_code == 400
    assert connected.post("/api/snapshots", json={"name": "../escape"}).status_code == 400
    assert connected.post("/api/snapshots", json={"name": "baseline"}).status_code == 201
    assert connected.post("/api/snapshots", json={"name": "baseline"}).status_code == 409


def _snapshot_bytes(plc_name=PLC_NAME, aois=None, created_at=None):
    data = snapshots.build_snapshot(
        plc_name,
        aois if aois is not None else AOIS,
        "4.10.6",
        created_at=created_at or T0,
    )
    return json.dumps(data).encode("utf-8")


def test_upload_snapshot(connected):
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("TestPLC_20261124-100000.json", _snapshot_bytes(), "application/json")},
    )
    assert response.status_code == 201, response.json()
    assert response.json()["name"] == "TestPLC_20261124-100000"
    entries = connected.get("/api/snapshots").json()
    assert [entry["name"] for entry in entries] == ["TestPLC_20261124-100000"]


def test_upload_snapshot_wrong_plc_409(connected):
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("OtherPLC.json", _snapshot_bytes(plc_name="OtherPLC"), "application/json")},
    )
    assert response.status_code == 409
    assert connected.get("/api/snapshots").json() == []


def test_upload_snapshot_rejects_bad_files(connected):
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("snap.txt", _snapshot_bytes(), "text/plain")},
    )
    assert response.status_code == 400
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("snap.json", b"{not json", "application/json")},
    )
    assert response.status_code == 400
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("snap.json", json.dumps({"format": "nope"}).encode(), "application/json")},
    )
    assert response.status_code == 400
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("big.json", b" " * (snapshots.MAX_UPLOAD_BYTES + 1), "application/json")},
    )
    assert response.status_code == 400


def test_restore_skips_invalid_values(connected):
    bad_aois = {
        "P_AIn": {
            "FIT_101": {
                ".Cfg_Desc": "x" * 41,  # over the STRING_40 limit
                ".Cfg_HiLim": 77.0,  # valid
                ".HiHi.Cfg_Exists": "maybe",  # not a BOOL
            }
        }
    }
    response = connected.post(
        "/api/snapshots/upload",
        files={"file": ("bad-values.json", _snapshot_bytes(aois=bad_aois), "application/json")},
    )
    assert response.status_code == 201
    response = connected.post("/api/snapshots/bad-values/restore")
    assert response.status_code == 200
    body = response.json()
    assert body["staged"] == 1
    assert [(entry["subtag"], entry["reason"]) for entry in body["skipped"]] == [
        (".Cfg_Desc", "string is 41 characters, longer than the 40-character limit for STRING_40"),
        (".HiHi.Cfg_Exists", "not a valid BOOL (expected 0/1/true/false), got 'maybe'"),
    ]
    staged = connected.get("/api/changes").json()
    assert [c["tag"] for c in staged] == ["FIT_101.Cfg_HiLim"]
    assert staged[0]["new_value"] == 77.0


def test_restore_refuses_other_plc_snapshot(connected, tmp_path):
    other = snapshots.build_snapshot("OtherPLC", AOIS, "4.10.6", created_at=T0)
    snapshots.save_snapshot(tmp_path / "snapshots", other, name="other-plc")
    assert connected.post("/api/snapshots/other-plc/restore").status_code == 409
    assert connected.get("/api/changes").json() == []


def test_write_takes_auto_snapshot(connected):
    connected.post(
        "/api/aois/P_AIn/paste",
        json={"tsv": "95", "anchor_instance": "FIT_101", "anchor_subtag": ".Cfg_HiLim"},
    )
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 0, "failed": 0}
    auto = body["auto_snapshot"]
    assert auto["ok"] is True
    assert auto["name"].startswith("TestPLC_")
    entries = connected.get("/api/snapshots").json()
    assert len(entries) == 1
    assert entries[0]["auto"] is True
    assert entries[0]["name"] == auto["name"]
    writes = [e for e in connected.get("/api/audit").json() if e["action"] == "write"]
    assert writes[0]["auto_snapshot"]["ok"] is True


def test_auto_snapshots_prune_to_newest_10(connected):
    for i in range(11):
        connected.post(
            "/api/aois/P_AIn/paste",
            json={"tsv": str(91 + i), "anchor_instance": "FIT_101", "anchor_subtag": ".Cfg_HiLim"},
        )
        body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
        assert body["auto_snapshot"]["ok"] is True
    entries = connected.get("/api/snapshots").json()
    assert len(entries) == 10
    assert all(entry["auto"] for entry in entries)


def test_auto_snapshot_captures_fresh_pre_write_values(connected, app):
    """F2: the auto-snapshot is the pre-write state, not the stale tab baseline.

    The tab was loaded (baseline 90.0), then the PLC drifted to 77.0.
    The auto-snapshot taken before the write must hold 77.0 — the
    value the write is about to leave behind — not the cached 90.0.
    A stale auto-snapshot would restore the drift away.
    """
    assert connected.get("/api/aois/P_AIn").status_code == 200
    app.state.pypax.plc._values["FIT_101.Cfg_HiLim"] = 77.0
    connected.post(
        "/api/aois/P_AIn/paste",
        json={
            "tsv": "New desc",
            "anchor_instance": "FIT_101",
            "anchor_subtag": ".Cfg_Desc",
        },
    )
    body = connected.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["auto_snapshot"]["ok"] is True
    auto_path = app.state.pypax.snapshot_dir / "auto" / f"{body['auto_snapshot']['name']}.json"
    saved = snapshots.load_snapshot(auto_path)
    assert saved["aois"]["P_AIn"]["FIT_101"][".Cfg_HiLim"] == 77.0


def test_corrupt_snapshot_can_be_deleted_and_downloaded(connected, tmp_path):
    """F3: a corrupt snapshot file must not get stuck in the store.

    The listing skips it (it cannot be parsed), but download serves
    the raw bytes for inspection and delete removes it — compare and
    restore still refuse it with 400 while it is there.
    """
    (tmp_path / "snapshots").mkdir(parents=True, exist_ok=True)
    (tmp_path / "snapshots" / "corrupt.json").write_text("{not json", encoding="utf-8")
    assert connected.get("/api/snapshots").json() == []
    assert connected.post("/api/snapshots/corrupt/compare").status_code == 400
    assert connected.post("/api/snapshots/corrupt/restore").status_code == 400
    response = connected.get("/api/snapshots/corrupt")
    assert response.status_code == 200
    assert response.content == b"{not json"
    response = connected.delete("/api/snapshots/corrupt")
    assert response.status_code == 200
    assert response.json() == {"removed": "corrupt"}
    assert not (tmp_path / "snapshots" / "corrupt.json").exists()


def test_compare_response_names_the_snapshot_plc(connected, tmp_path):
    """Compare of a snapshot for another PLC names that PLC, so the UI
    can warn instead of showing a wall of diffs with no context."""
    other = snapshots.build_snapshot("OtherPLC", AOIS, "4.10.6", created_at=T0)
    snapshots.save_snapshot(tmp_path / "snapshots", other, name="other-plc")
    body = connected.post("/api/snapshots/other-plc/compare").json()
    assert body["plc_name"] == "OtherPLC"
    assert body["differences"]


_L5X_STRING_CUSTOM = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<RSLogix5000Content SchemaRevision="1.0" TargetName="P_StrTest" '
    'TargetType="AddOnInstructionDefinition" TargetRevision="1.00 Test">\n'
    "<DataTypes>\n"
    '<DataType Name="STRING_Custom" Family="StringFamily" Class="User">\n'
    "<Members>\n"
    '<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal"/>\n'
    '<Member Name="DATA" DataType="SINT" Dimension="12" Radix="ASCII"/>\n'
    "</Members>\n"
    "</DataType>\n"
    "</DataTypes>\n"
    '<Controller Use="Context" Name="Test">\n'
    '<AddOnInstructionDefinition Name="P_StrTest" Revision="1.00" Vendor="Test">\n'
    "<Parameters>\n"
    '<Parameter Name="Cfg_Label" TagType="Base" DataType="STRING_Custom" '
    'Usage="Input" Radix="Decimal" Required="false" Visible="false" '
    'ExternalAccess="Read/Write">\n'
    "<Description><![CDATA[Label]]></Description>\n"
    "</Parameter>\n"
    "</Parameters>\n"
    "</AddOnInstructionDefinition></Controller></RSLogix5000Content>\n"
)


def test_restore_uses_configured_string_registry(workbook, tmp_path):
    """F5: restore validates strings against the app's registry, like
    paste does — not the repo-default registry file. A 20-character
    value for a STRING_Custom column (limit 12 in the configured
    registry) must be skipped by restore exactly as paste rejects it.
    """
    registry = tmp_path / "string_udt_sizes.yaml"
    registry.write_text("STRING_CUSTOM:\n  size: 12\n  source: probe.L5X\n", encoding="utf-8")

    def factory(commpath):
        plc = FakePLC(commpath)
        plc.add_instance("STR_101", "P_StrTest", [(".Cfg_Label", "STRING", "short")])
        return plc

    app = create_app(
        plc_factory=factory,
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
        snapshot_dir=str(tmp_path / "snapshots"),
        string_registry_path=str(registry),
    )
    client = TestClient(app)
    assert client.post("/api/connect", json={"commpath": "10.0.0.1/0"}).json()["connected"]
    upload = client.post(
        "/api/aois/import", files={"file": ("P_StrTest.L5X", _L5X_STRING_CUSTOM.encode())}
    )
    assert upload.status_code == 200
    confirm = client.post(
        "/api/aois/import/confirm", json={"proposal_id": upload.json()["proposal_id"]}
    )
    assert confirm.status_code == 200
    # The paste path honours the configured registry: 20 chars > 12.
    paste = client.post(
        "/api/aois/P_StrTest/paste",
        json={
            "tsv": "x" * 20,
            "anchor_instance": "STR_101",
            "anchor_subtag": ".Cfg_Label",
        },
    )
    assert len(paste.json()["rejected"]) == 1
    assert "12-character limit" in paste.json()["rejected"][0]["reason"]
    # A snapshot holding the same over-long value must be skipped, not staged.
    snap = snapshots.build_snapshot(
        PLC_NAME,
        {"P_StrTest": {"STR_101": {".Cfg_Label": "x" * 20}}},
        "4.10.6",
        created_at=T0,
    )
    response = client.post(
        "/api/snapshots/upload",
        files={"file": ("long.json", json.dumps(snap).encode(), "application/json")},
    )
    assert response.status_code == 201
    body = client.post("/api/snapshots/long/restore").json()
    assert body["staged"] == 0
    assert len(body["skipped"]) == 1
    assert "12-character limit" in body["skipped"][0]["reason"]


def test_failed_auto_snapshot_never_blocks_write(workbook, tmp_path):
    blocker = tmp_path / "not-a-dir"
    blocker.write_text("a file, not a directory", encoding="utf-8")
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
        snapshot_dir=str(blocker),
    )
    client = TestClient(app)
    assert client.post("/api/connect", json={"commpath": "10.0.0.1/0"}).json()["connected"]
    client.post(
        "/api/aois/P_AIn/paste",
        json={"tsv": "95", "anchor_instance": "FIT_101", "anchor_subtag": ".Cfg_HiLim"},
    )
    body = client.post("/api/write", json={"plc_name": PLC_NAME}).json()
    assert body["summary"] == {"verified": 1, "mismatch": 0, "failed": 0}
    assert body["auto_snapshot"]["ok"] is False
    assert body["auto_snapshot"]["error"]
    writes = [e for e in client.get("/api/audit").json() if e["action"] == "write"]
    assert writes[0]["auto_snapshot"]["ok"] is False
    assert writes[0]["results"] == {"verified": 1, "mismatch": 0, "failed": 0}


# ----------------------------------------------------------------------
# Dev mock server over real HTTP (scripts/dev_mock_api.py)
# ----------------------------------------------------------------------


def _load_dev_mock_api():
    import importlib.util

    root = Path(__file__).resolve().parents[1]
    spec = importlib.util.spec_from_file_location(
        "dev_mock_api_snapshots", root / "scripts" / "dev_mock_api.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_mock_server_snapshot_flow_over_http():
    """The dev mock server answers the snapshot contract end to end."""
    import http.client
    import threading

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

        status, body = call("GET", "/api/snapshots")
        assert status == 200 and body == []

        status, body = call("POST", "/api/snapshots", {"name": "baseline"})
        assert status == 201 and body["name"] == "baseline"
        assert body["plc_name"] == "MOCK_PLC"

        # Move the mock PLC (paste + write), then compare with the snapshot.
        call(
            "POST",
            "/api/aois/P_AIn/paste",
            {"tsv": "95", "anchor_instance": "FIT-101", "anchor_subtag": ".Cfg_HiLim"},
        )
        status, body = call("POST", "/api/write", {"plc_name": "MOCK_PLC"})
        assert status == 200
        assert body["auto_snapshot"]["ok"] is True

        status, body = call("POST", "/api/snapshots/baseline/compare")
        assert status == 200
        assert body["differences"] == [
            {
                "aoi": "P_AIn",
                "instance": "FIT-101",
                "subtag": ".Cfg_HiLim",
                "snapshot_value": 90,
                "current_value": 95.0,
            }
        ]

        status, body = call("POST", "/api/snapshots/baseline/restore")
        assert status == 200
        assert body == {"staged": 1, "skipped": []}
        status, changes = call("GET", "/api/changes")
        assert len(changes) == 1 and changes[0]["new_value"] == 90.0

        status, body = call("POST", "/api/write", {"plc_name": "MOCK_PLC"})
        assert status == 200 and body["summary"]["verified"] == 1

        # Two auto-snapshots (one per write) plus the manual one.
        status, entries = call("GET", "/api/snapshots")
        assert status == 200 and len(entries) == 3
        assert sum(1 for e in entries if e["auto"]) == 2

        conn.request("GET", "/api/snapshots/baseline")
        res = conn.getresponse()
        payload = json.loads(res.read().decode("utf-8"))
        assert res.status == 200
        assert payload["format"] == "pypax-snapshot/1"
        assert "attachment" in (res.getheader("Content-Disposition") or "")

        status, body = call("DELETE", "/api/snapshots/baseline")
        assert status == 200 and body == {"removed": "baseline"}
        status, entries = call("GET", "/api/snapshots")
        assert len(entries) == 2
        conn.close()
    finally:
        server.shutdown()
        server.server_close()
