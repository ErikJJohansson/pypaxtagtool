"""Tests for re-importing a custom AOI in place (Phase 9c).

Upload or scan a changed ``.L5X`` for an already-imported AOI and the
backend answers with ``status: "update_available"`` and a column diff
instead of a 409; confirming the update rewrites
``config/aoi/custom/<AOI>.yaml`` in place. Confirming an update while
staged changes exist on the AOI is refused (409). Change detection is
a SHA-256 of the source file stored in the YAML header at confirm
time, with a column-diff fallback for YAMLs written before hashes
existed. The ``.L5X`` fixtures are small synthetic exports built
in-test — never a real Rockwell file.
"""

from __future__ import annotations

import hashlib
import json
import struct

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax.custom_aoi import diff_columns, hash_file, read_source_sha256, save_custom_aoi
from pypax.l5x import Column
from pypax.web import create_app

PLC_NAME = "TestPLC"


def _member_xml(name, data_type, description, default):
    """One ``<Parameter>``: default is ("hex", bytes) or ("str", text)."""
    if default is None:
        default_xml = ""
    elif default[0] == "hex":
        default_xml = f"<DefaultData>{default[1]}</DefaultData>"
    else:
        default_xml = f"<DefaultData Format=\"String\"><![CDATA['{default[1]}']]></DefaultData>"
    return (
        f'<Parameter Name="{name}" TagType="Base" DataType="{data_type}" '
        f'Usage="Input" Radix="Decimal" Required="false" Visible="false" '
        f'ExternalAccess="Read/Write">'
        f"<Description><![CDATA[{description}]]></Description>"
        f"{default_xml}</Parameter>"
    )


def _l5x(aoi_name, members, revision="1.02"):
    """A minimal synthetic ``.L5X`` export for ``aoi_name``."""
    params = "".join(_member_xml(*member) for member in members)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<RSLogix5000Content SchemaRevision="1.0" TargetName="{aoi_name}" '
        f'TargetType="AddOnInstructionDefinition" TargetRevision="{revision}.00 Test">'
        f'<Controller Use="Context" Name="Test">'
        f'<AddOnInstructionDefinition Name="{aoi_name}" Revision="{revision}" '
        'Vendor="Test">'
        f"<Parameters>{params}</Parameters>"
        "</AddOnInstructionDefinition></Controller></RSLogix5000Content>"
    )


def _dint(value):
    return ("hex", struct.pack("<i", value).hex())


def _real(value):
    return ("hex", struct.pack("<f", value).hex())


V1_MEMBERS = [
    ("Cfg_Desc", "STRING_40", "Description", ("str", "Valve")),
    ("Cfg_HasStop", "BOOL", "Has stop pushbutton", ("hex", "00")),
    ("Cfg_Total", "REAL", "Total", _real(5.0)),
    ("Cfg_Count", "DINT", "Count", _dint(42)),
]
V1_L5X = _l5x("P_TestValve", V1_MEMBERS)
V1_SUBTAGS = [".Cfg_Desc", ".Cfg_HasStop", ".Cfg_Total", ".Cfg_Count"]

# v2: description changed, type changed (REAL -> DINT, default 5.0 -> 5),
# one column added, one removed, one untouched.
V2_MEMBERS = [
    ("Cfg_Desc", "STRING_40", "Description for display", ("str", "Valve")),
    ("Cfg_HasStop", "BOOL", "Has stop pushbutton", ("hex", "00")),
    ("Cfg_Total", "DINT", "Total", _dint(5)),
    ("Cfg_New", "REAL", "New column", _real(1.5)),
]
V2_L5X = _l5x("P_TestValve", V2_MEMBERS)
V2_SUBTAGS = [".Cfg_Desc", ".Cfg_HasStop", ".Cfg_Total", ".Cfg_New"]

VALVE_SEED = [
    (".Cfg_Desc", "STRING", "Test Valve 101"),
    (".Cfg_HasStop", "BOOL", 0),
    (".Cfg_Total", "REAL", 5.0),
    (".Cfg_Count", "DINT", 42),
]


@pytest.fixture()
def workbook(tmp_path):
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
        plc_factory=_plc_factory,
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
    )


def _plc_factory(commpath):
    plc = FakePLC(commpath)
    plc.add_instance("VLV_101", "P_TestValve", VALVE_SEED)
    return plc


@pytest.fixture()
def client(app):
    return TestClient(app)


@pytest.fixture()
def connected(client):
    response = client.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    assert response.status_code == 200
    assert response.json()["connected"] is True
    return client


def _upload(client, content, filename="P_TestValve.L5X"):
    return client.post("/api/aois/import", files={"file": (filename, content)})


def _confirm(client, proposal_id, selected=None):
    body = {"proposal_id": proposal_id}
    if selected is not None:
        body["selected"] = selected
    return client.post("/api/aois/import/confirm", json=body)


def _import_v1(client):
    """Upload + confirm the v1 export; returns the confirm response."""
    proposal = _upload(client, V1_L5X.encode()).json()
    assert proposal["aoi"] == "P_TestValve"
    assert [c["subtag"] for c in proposal["columns"]] == V1_SUBTAGS
    response = _confirm(client, proposal["proposal_id"])
    assert response.status_code == 200
    return response


def _changed_by_subtag(diff):
    return {entry["subtag"]: entry for entry in diff["changed"]}


# --- Store helpers: hash + diff -------------------------------------------


def test_hash_recorded_in_yaml_header_and_read_back(tmp_path):
    target = save_custom_aoi(
        "P_TestValve",
        [Column(subtag=".Cfg_Desc", data_type="STRING_40")],
        "P_TestValve.L5X",
        "1.02",
        path=tmp_path,
        source_sha256="a" * 64,
    )
    assert "# Source SHA-256: " + "a" * 64 in target.read_text(encoding="utf-8")
    assert read_source_sha256(target) == "a" * 64
    # A YAML written without a hash, and a missing file, read as None.
    plain = save_custom_aoi(
        "P_Other",
        [Column(subtag=".Cfg_Desc", data_type="STRING_40")],
        "P_Other.L5X",
        "1.02",
        path=tmp_path,
    )
    assert read_source_sha256(plain) is None
    assert read_source_sha256(tmp_path / "nope.yaml") is None


def test_hash_file_matches_sha256(tmp_path):
    path = tmp_path / "P_TestValve.L5X"
    path.write_bytes(V1_L5X.encode())
    assert hash_file(path) == hashlib.sha256(V1_L5X.encode()).hexdigest()


def test_diff_columns():
    before = {
        ".Cfg_Desc": Column(
            subtag=".Cfg_Desc", data_type="STRING_40", description="Desc", default="Valve"
        ),
        ".Cfg_Keep": Column(subtag=".Cfg_Keep", data_type="BOOL", description="", default=False),
        # Details unknown (source .L5X gone): counts as unchanged.
        ".Cfg_Mystery": None,
        ".Cfg_Gone": Column(subtag=".Cfg_Gone", data_type="DINT", description="", default=1),
    }
    after = [
        Column(subtag=".Cfg_Desc", data_type="STRING_40", description="New desc", default="Valve"),
        Column(subtag=".Cfg_Keep", data_type="BOOL", description="", default=False),
        Column(subtag=".Cfg_Mystery", data_type="REAL", description="Changed?", default=1.5),
        Column(subtag=".Cfg_Added", data_type="DINT", description="", default=0),
    ]
    diff = diff_columns(before, after)
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_Added"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Gone"]
    assert [c["subtag"] for c in diff["changed"]] == [".Cfg_Desc"]
    changed = diff["changed"][0]
    assert changed["before"]["description"] == "Desc"
    assert changed["after"]["description"] == "New desc"
    # .Cfg_Keep and the unknown-details .Cfg_Mystery are unchanged.
    assert diff["unchanged"] == 2


# --- Upload re-import ------------------------------------------------------


def test_reimport_unchanged_file_empty_diff(client):
    _import_v1(client)
    second = _upload(client, V1_L5X.encode()).json()
    assert second["status"] == "update_available"
    assert second["update_available"] is True
    assert second["proposal_id"]
    assert second["diff"] == {
        "added": [],
        "removed": [],
        "changed": [],
        "unchanged": 4,
        "before_details_known": True,
    }


def test_reimport_changed_file_correct_diff(client):
    _import_v1(client)
    second = _upload(client, V2_L5X.encode()).json()
    assert second["status"] == "update_available"
    diff = second["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Count"]
    changed = _changed_by_subtag(diff)
    assert set(changed) == {".Cfg_Desc", ".Cfg_Total"}
    assert changed[".Cfg_Desc"]["before"]["description"] == "Description"
    assert changed[".Cfg_Desc"]["after"]["description"] == "Description for display"
    assert changed[".Cfg_Total"]["before"] == {
        "subtag": ".Cfg_Total",
        "type": "REAL",
        "description": "Total",
        "default": 5.0,
    }
    assert changed[".Cfg_Total"]["after"] == {
        "subtag": ".Cfg_Total",
        "type": "DINT",
        "description": "Total",
        "default": 5,
    }
    assert diff["unchanged"] == 1  # .Cfg_HasStop only


def test_upload_with_colliding_filename_but_different_aoi_is_refused(client, tmp_path):
    """F4: an upload must not silently overwrite another AOI's export.

    The upload is persisted to the import directory before confirm, so
    a file named ``P_A.L5X`` holding AOI ``P_B`` would destroy the
    ``P_A`` export on disk. Refuse it (400); the original file stays.
    """
    first = _upload(client, _l5x("P_A", V1_MEMBERS).encode(), filename="P_A.L5X")
    assert first.status_code == 200
    assert first.json()["aoi"] == "P_A"
    assert _confirm(client, first.json()["proposal_id"]).status_code == 200
    response = _upload(client, _l5x("P_B", V1_MEMBERS).encode(), filename="P_A.L5X")
    assert response.status_code == 400
    detail = response.json()["detail"]
    assert "P_B" in detail and "different AOI" in detail
    on_disk = (tmp_path / "imports" / "P_A.L5X").read_text(encoding="utf-8")
    assert 'TargetName="P_A"' in on_disk
    # The same AOI under the same filename is a legitimate re-import.
    again = _upload(client, _l5x("P_A", V2_MEMBERS).encode(), filename="P_A.L5X")
    assert again.status_code == 200
    assert again.json()["status"] == "update_available"


def test_confirm_update_rewrites_yaml(client, tmp_path):
    _import_v1(client)
    yaml_path = tmp_path / "custom" / "P_TestValve.yaml"
    assert "# Source SHA-256: " + hashlib.sha256(V1_L5X.encode()).hexdigest() in (
        yaml_path.read_text(encoding="utf-8")
    )
    second = _upload(client, V2_L5X.encode()).json()
    response = _confirm(client, second["proposal_id"])
    assert response.status_code == 200
    assert response.json() == {"aoi": "P_TestValve", "columns": 4, "updated": True}
    text = yaml_path.read_text(encoding="utf-8")
    for subtag in V2_SUBTAGS:
        assert f"  - '{subtag}'" in text
    assert "  - '.Cfg_Count'" not in text
    assert "# Source SHA-256: " + hashlib.sha256(V2_L5X.encode()).hexdigest() in text
    # The proposal is consumed: confirming the same id again is unknown.
    assert _confirm(client, second["proposal_id"]).status_code == 404


def test_confirm_update_rewrites_model(connected):
    _import_v1(connected)
    second = _upload(connected, V2_L5X.encode()).json()
    assert _confirm(connected, second["proposal_id"]).status_code == 200
    aois = connected.get("/api/aois").json()
    assert {
        "name": "P_TestValve",
        "instance_count": 1,
        "subtag_count": 4,
        "custom": True,
    } in aois
    body = connected.get("/api/aois/P_TestValve").json()
    assert [c["subtag"] for c in body["columns"]] == V2_SUBTAGS
    by_subtag = {c["subtag"]: c for c in body["columns"]}
    assert by_subtag[".Cfg_Total"]["data_type"] == "DINT"
    assert by_subtag[".Cfg_New"]["data_type"] == "REAL"


def test_confirm_update_ignores_selected(client):
    """An update is all-or-nothing: a selection does not filter it."""
    _import_v1(client)
    second = _upload(client, V2_L5X.encode()).json()
    response = _confirm(client, second["proposal_id"], selected=[])
    assert response.status_code == 200
    assert response.json() == {"aoi": "P_TestValve", "columns": 4, "updated": True}


def test_confirm_update_with_staged_changes_is_409(connected, tmp_path):
    _import_v1(connected)
    staged = connected.post(
        "/api/aois/P_TestValve/paste",
        json={"tsv": "7.5\n", "anchor_instance": "VLV_101", "anchor_subtag": ".Cfg_Total"},
    )
    assert staged.status_code == 200
    assert staged.json()["applied"] == 1
    second = _upload(connected, V2_L5X.encode()).json()
    response = _confirm(connected, second["proposal_id"])
    assert response.status_code == 409
    assert "staged changes" in response.json()["detail"]
    # The staged edit survives, the YAML is untouched, the proposal lives.
    assert len(connected.get("/api/changes").json()) == 1
    text = (tmp_path / "custom" / "P_TestValve.yaml").read_text(encoding="utf-8")
    assert "  - '.Cfg_Count'" in text
    assert "  - '.Cfg_New'" not in text
    # Discard the staged change, then the same proposal confirms fine.
    assert connected.delete("/api/changes").status_code == 200
    assert _confirm(connected, second["proposal_id"]).status_code == 200


def test_staged_guard_is_case_insensitive(connected):
    """The staged-change guard cannot be bypassed by re-importing the
    AOI under a differently-cased name."""
    _import_v1(connected)
    staged = connected.post(
        "/api/aois/P_TestValve/paste",
        json={"tsv": "7.5\n", "anchor_instance": "VLV_101", "anchor_subtag": ".Cfg_Total"},
    )
    assert staged.json()["applied"] == 1
    lower = _l5x("p_testvalve", V2_MEMBERS)
    second = _upload(connected, lower.encode(), filename="p_testvalve.L5X").json()
    assert second["status"] == "update_available"
    assert second["aoi"] == "p_testvalve"
    response = _confirm(connected, second["proposal_id"])
    assert response.status_code == 409
    assert "staged changes" in response.json()["detail"]


def test_case_insensitive_name_match(client, tmp_path):
    _import_v1(client)
    lower = _l5x("p_testvalve", V2_MEMBERS)
    second = _upload(client, lower.encode(), filename="p_testvalve.L5X").json()
    assert second["status"] == "update_available"
    assert _confirm(client, second["proposal_id"]).status_code == 200
    # Still one file, under the canonical name — no p_testvalve.yaml.
    assert sorted(p.name for p in (tmp_path / "custom").glob("*.yaml")) == ["P_TestValve.yaml"]
    from pypax.custom_aoi import load_custom_aois

    assert load_custom_aois(tmp_path / "custom") == {"P_TestValve": V2_SUBTAGS}


def test_update_is_audited(client):
    _import_v1(client)
    second = _upload(client, V2_L5X.encode()).json()
    assert _confirm(client, second["proposal_id"]).status_code == 200
    entries = {e["action"]: e for e in client.get("/api/audit").json()}
    assert entries["update_custom_aoi"]["results"] == {
        "aoi": "P_TestValve",
        "columns": 4,
        "added": 1,
        "removed": 1,
        "changed": 2,
    }


# --- Folder scan -----------------------------------------------------------


def test_scan_detects_changed_file(client, tmp_path):
    _import_v1(client)
    # Unchanged file: still skipped as already imported, no update.
    body = client.get("/api/aois/scan").json()
    assert body["proposals"] == []
    assert body["updates"] == []
    assert body["skipped"] == [{"file": "P_TestValve.L5X", "reason": "already imported"}]
    # The export changes on disk: the scan reports an update with a diff.
    (tmp_path / "imports" / "P_TestValve.L5X").write_bytes(V2_L5X.encode())
    body = client.get("/api/aois/scan").json()
    assert body["proposals"] == []
    assert body["skipped"] == []
    assert len(body["updates"]) == 1
    update = body["updates"][0]
    assert update["aoi"] == "P_TestValve"
    assert update["status"] == "update_available"
    assert update["update_available"] is True
    assert update["source_file"] == "P_TestValve.L5X"
    assert update["proposal_id"]
    assert [c["subtag"] for c in update["diff"]["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in update["diff"]["removed"]] == [".Cfg_Count"]
    assert update["diff"]["unchanged"] == 1
    # Confirming the scanned update works, and the next scan is quiet.
    assert _confirm(client, update["proposal_id"]).status_code == 200
    body = client.get("/api/aois/scan").json()
    assert body["updates"] == []
    assert body["skipped"] == [{"file": "P_TestValve.L5X", "reason": "already imported"}]


def test_scan_without_stored_hash_falls_back_to_diff(client, tmp_path):
    """A YAML written before hashes existed: the column diff decides."""
    from pypax.custom_aoi import save_custom_aoi
    from pypax.l5x import Column as L5xColumn

    imports = tmp_path / "imports"
    imports.mkdir()
    (imports / "P_TestValve.L5X").write_bytes(V2_L5X.encode())
    # Legacy YAML: v1 columns, no hash line.
    save_custom_aoi(
        "P_TestValve",
        [L5xColumn(subtag=s, data_type=None) for s in V1_SUBTAGS],
        "P_TestValve.L5X",
        "1.02",
        path=tmp_path / "custom",
    )
    body = client.get("/api/aois/scan").json()
    assert len(body["updates"]) == 1
    assert body["updates"][0]["aoi"] == "P_TestValve"
    diff = body["updates"][0]["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Count"]
    # No column metadata survives for the legacy import, so the common
    # columns cannot be compared for type/description/default changes.
    assert diff["changed"] == []
    assert diff["unchanged"] == 3
    assert diff["before_details_known"] is False


def test_second_unconfirmed_upload_still_diffs_against_confirmed(client):
    """A superseded unconfirmed proposal must not move the "before" side.

    v2 is proposed but never confirmed; v3 is uploaded next. The diff
    is still confirmed-v1 vs v3 — the unconfirmed v2 proposal must not
    become the baseline.
    """
    _import_v1(client)
    assert _upload(client, V2_L5X.encode()).json()["status"] == "update_available"
    v3_members = [
        ("Cfg_Desc", "STRING_40", "Third description", ("str", "Valve")),
        ("Cfg_HasStop", "BOOL", "Has stop pushbutton", ("hex", "00")),
        ("Cfg_Total", "DINT", "Total", _dint(5)),
        ("Cfg_New", "REAL", "New column", _real(1.5)),
    ]
    third = _upload(client, _l5x("P_TestValve", v3_members).encode()).json()
    assert third["status"] == "update_available"
    diff = third["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Count"]
    changed = _changed_by_subtag(diff)
    assert changed[".Cfg_Desc"]["before"]["description"] == "Description"
    assert changed[".Cfg_Desc"]["after"]["description"] == "Third description"
    assert changed[".Cfg_Total"]["before"]["type"] == "REAL"
    assert changed[".Cfg_Total"]["after"]["type"] == "DINT"
    assert diff["unchanged"] == 1


def _restarted_client(workbook, tmp_path):
    """A fresh app against the same dirs (simulates an app restart)."""
    fresh = create_app(
        plc_factory=_plc_factory,
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
    )
    return TestClient(fresh)


def test_restart_with_unchanged_file_diffs_correctly(client, workbook, tmp_path):
    """Startup seeds the confirmed cache when the file hash matches."""
    _import_v1(client)
    restarted = _restarted_client(workbook, tmp_path)
    second = _upload(restarted, V2_L5X.encode()).json()
    assert second["status"] == "update_available"
    diff = second["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Count"]
    assert set(_changed_by_subtag(diff)) == {".Cfg_Desc", ".Cfg_Total"}
    assert diff["unchanged"] == 1


def test_sidecar_written_on_confirm_and_removed_on_delete(client, tmp_path):
    """F6: confirming writes the confirmed-metadata sidecar next to the
    YAML, and deleting the custom AOI removes it with the YAML."""
    _import_v1(client)
    sidecar = tmp_path / "custom" / "P_TestValve.meta.json"
    assert sidecar.is_file()
    meta = json.loads(sidecar.read_text(encoding="utf-8"))
    assert meta[".Cfg_Desc"] == {
        "type": "STRING_40",
        "description": "Description",
        "default": "Valve",
    }
    assert meta[".Cfg_Total"] == {"type": "REAL", "description": "Total", "default": 5.0}
    assert client.delete("/api/aois/custom/P_TestValve").status_code == 200
    assert not sidecar.exists()
    assert not (tmp_path / "custom" / "P_TestValve.yaml").exists()


def test_restart_with_changed_file_still_diffs_changed(client, workbook, tmp_path):
    """F6: after a restart with a changed file on disk, the re-import
    diff still detects type/description/default changes.

    The confirmed column details are persisted next to the YAML at
    confirm time (a sidecar), so the "before" side of the diff does
    not depend on the in-memory cache or on the source file still
    matching its recorded hash.
    """
    _import_v1(client)
    (tmp_path / "imports" / "P_TestValve.L5X").write_bytes(V2_L5X.encode())
    restarted = _restarted_client(workbook, tmp_path)
    body = restarted.get("/api/aois/scan").json()
    assert len(body["updates"]) == 1
    diff = body["updates"][0]["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert [c["subtag"] for c in diff["removed"]] == [".Cfg_Count"]
    assert set(_changed_by_subtag(diff)) == {".Cfg_Desc", ".Cfg_Total"}
    assert diff["unchanged"] == 1
    assert diff["before_details_known"] is True


def test_restart_with_deleted_source_file_still_diffs_changed(client, workbook, tmp_path):
    """F6: after a restart with the source .L5X deleted from the import
    folder, uploading the new version still diffs against the confirmed
    columns (the sidecar survives the missing file)."""
    _import_v1(client)
    (tmp_path / "imports" / "P_TestValve.L5X").unlink()
    restarted = _restarted_client(workbook, tmp_path)
    second = _upload(restarted, V2_L5X.encode()).json()
    assert second["status"] == "update_available"
    diff = second["diff"]
    assert set(_changed_by_subtag(diff)) == {".Cfg_Desc", ".Cfg_Total"}
    assert diff["before_details_known"] is True


def test_legacy_yaml_without_sidecar_reports_unknown_details(client, workbook, tmp_path):
    """A custom AOI written before sidecars existed (hand-written YAML,
    no confirmed metadata anywhere) diffs added/removed only — and the
    diff says so, instead of reporting unchanged columns it cannot see."""
    custom_dir = tmp_path / "custom"
    custom_dir.mkdir(parents=True, exist_ok=True)
    (custom_dir / "P_Legacy.yaml").write_text(
        "# Custom AOI subtag definitions.\nP_Legacy:\n  - '.Cfg_Desc'\n  - '.Cfg_Total'\n",
        encoding="utf-8",
    )
    v2_members = [
        ("Cfg_Desc", "STRING_40", "Description v2", ("str", "Valve")),
        ("Cfg_Total", "DINT", "Total", _dint(5)),
        ("Cfg_New", "REAL", "New column", _real(1.5)),
    ]
    second = _upload(client, _l5x("P_Legacy", v2_members).encode()).json()
    assert second["status"] == "update_available"
    diff = second["diff"]
    assert [c["subtag"] for c in diff["added"]] == [".Cfg_New"]
    assert diff["removed"] == []
    assert diff["changed"] == []
    assert diff["before_details_known"] is False


# --- Frontend surface ------------------------------------------------------


def test_frontend_update_tokens():
    """The scan list and modal carry the re-import UI (Phase 9c)."""
    from pathlib import Path

    root = Path(__file__).resolve().parents[1]
    app_js = (root / "src" / "pypax" / "web" / "static" / "app.js").read_text(encoding="utf-8")
    styles = (root / "src" / "pypax" / "web" / "static" / "styles.css").read_text(encoding="utf-8")
    for token in (
        "Update available",
        "Update AOI",
        "update_available",
        "diffSummary",
        "renderImportDiff",
        "diff-kind",
        "MOCK_PROPOSAL_V2",
    ):
        assert token in app_js, token
    for token in ("update-badge", "scan-update", "diff-added", "diff-removed", "diff-changed"):
        assert token in styles, token
