"""Tests for the custom-AOI import endpoints (Phase 8b).

Upload, folder scan, confirm, and delete against the FastAPI app
with a FakePLC. The ``.L5X`` fixture is a small synthetic export
built in-test — never a real Rockwell file.
"""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import build_workbook
from pypax.config import AppSettings
from pypax.web import create_app

PLC_NAME = "TestPLC"

_L5X = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="18.02" TargetName="P_TestValve" TargetType="AddOnInstructionDefinition" TargetRevision="1.02.00 Test" TargetLastEdited="2026-11-24T00:00:00Z" ContainsContext="true" Owner="Test" ExportDate="Mon Nov 24 00:00:00 2026" ExportOptions="References DecoratedData Context Dependencies AllProjDocTrans">
<Controller Use="Context" Name="Test">
<AddOnInstructionDefinition Name="P_TestValve" Revision="1.02" RevisionExtension=".00 Test" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Inp_OpenLS" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Open limit switch input]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_HasStop" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Has stop pushbutton]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_CoastToLS" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Coasting into LS when stopped]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Cfg_MyAlarmSeverity" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Pass-through: use MyAlarm.Cfg_Severity]]></Description>
<DefaultData>EE 02 00 00</DefaultData>
</Parameter>
<Parameter Name="Cfg_Total" TagType="Base" DataType="REAL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[A real-valued config]]></Description>
<DefaultData>00 00 A0 40</DefaultData>
</Parameter>
<Parameter Name="Cfg_Count" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[A DINT config]]></Description>
<DefaultData>2A 00 00 00</DefaultData>
</Parameter>
<Parameter Name="Sts_Opened" TagType="Base" DataType="BOOL" Usage="Output" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Opened status]]></Description>
</Parameter>
</Parameters>
<LocalTags>
<LocalTag Name="MyAlarm" DataType="P_Alarm" ExternalAccess="Read/Write">
<Description><![CDATA[My alarm object]]></Description>
</LocalTag>
<LocalTag Name="CmdSrc" DataType="P_CmdSrc" ExternalAccess="Read/Write">
<Description><![CDATA[Command source object]]></Description>
</LocalTag>
<LocalTag Name="Cfg_Tag" DataType="STRING_20" ExternalAccess="Read/Write">
<Description><![CDATA[Tagname for display on HMI]]></Description>
<DefaultData Format="String" Length="11"><![CDATA['P_TestValve']]></DefaultData>
</LocalTag>
<LocalTag Name="Cfg_Label" DataType="STRING_20" ExternalAccess="Read/Write">
<Description><![CDATA[Label]]></Description>
<DefaultData>0A 00 00 00 54 65 73 74 20 56 61 6C 76 65 00 00 00 00 00 00 00 00 00</DefaultData>
</LocalTag>
<LocalTag Name="Cfg_Desc" DataType="STRING_40" ExternalAccess="Read/Write">
<Description><![CDATA[Description for display on HMI]]></Description>
<DefaultData>0A 00 00 00 54 65 73 74 20 56 61 6C 76 65 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00</DefaultData>
</LocalTag>
<LocalTag Name="Cfg_Area" DataType="STRING_Area" ExternalAccess="Read/Write">
<Description><![CDATA[Process Area for security]]></Description>
<DefaultData>06 00 00 00 61 72 65 61 30 31 00 00</DefaultData>
</LocalTag>
<LocalTag Name="Wrk_Foo" DataType="BOOL" ExternalAccess="Read/Write">
<Description><![CDATA[Working tag]]></Description>
<DefaultData>00</DefaultData>
</LocalTag>
</LocalTags>
</AddOnInstructionDefinition>
<AddOnInstructionDefinition Name="P_Alarm" Revision="1.00" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Cfg_Exists" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Alarm exists]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Cfg_Severity" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Alarm severity]]></Description>
<DefaultData>EE 02 00 00</DefaultData>
</Parameter>
<Parameter Name="Sts_Err" TagType="Base" DataType="BOOL" Usage="Output" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Alarm error status]]></Description>
</Parameter>
</Parameters>
<LocalTags>
<LocalTag Name="Cfg_Tag" DataType="STRING_20" ExternalAccess="Read/Write">
<Description><![CDATA[Alarm Tag Text]]></Description>
<DefaultData Format="String"><![CDATA['P_Alarm']]></DefaultData>
</LocalTag>
<LocalTag Name="Cfg_Cond" DataType="STRING_20" ExternalAccess="Read/Write">
<Description><![CDATA[Alarm Condition Text]]></Description>
<DefaultData>0F 00 00 00 41 6C 61 72 6D 20 43 6F 6E 64 69 74 69 6F 6E 00 00 00 00 00</DefaultData>
</LocalTag>
</LocalTags>
</AddOnInstructionDefinition>
<AddOnInstructionDefinition Name="P_CmdSrc" Revision="1.00" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Cfg_HasOper" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Has operator]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Sts_Oper" TagType="Base" DataType="BOOL" Usage="Output" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Operator status]]></Description>
</Parameter>
</Parameters>
</AddOnInstructionDefinition>
</Controller>
</RSLogix5000Content>
"""

# Curated columns for the fixture, in proposal order.
EXPECTED_SUBTAGS = [
    ".Cfg_Desc",
    ".Cfg_Label",
    ".Cfg_Tag",
    ".Cfg_Area",
    ".Cfg_HasStop",
    ".Cfg_CoastToLS",
    ".Cfg_Total",
    ".Cfg_Count",
    ".MyAlarm.Cfg_Exists",
    ".MyAlarm.Cfg_Severity",
    ".MyAlarm.Cfg_Tag",
    ".MyAlarm.Cfg_Cond",
    ".CmdSrc.Cfg_HasOper",
]

# (subtag, type, value) seed for the FakePLC P_TestValve instance.
VALVE_SEED = [
    (".Cfg_Desc", "STRING", "Test Valve 101"),
    (".Cfg_Label", "STRING", "VLV-101"),
    (".Cfg_Tag", "STRING", "VLV_101"),
    (".Cfg_Area", "STRING", "area01"),
    (".Cfg_HasStop", "BOOL", 0),
    (".Cfg_CoastToLS", "BOOL", 1),
    (".Cfg_Total", "REAL", 5.0),
    (".Cfg_Count", "DINT", 42),
    (".MyAlarm.Cfg_Exists", "BOOL", 1),
    (".MyAlarm.Cfg_Severity", "DINT", 750),
    (".MyAlarm.Cfg_Tag", "STRING", "P_Alarm"),
    (".MyAlarm.Cfg_Cond", "STRING", "Alarm Condition"),
    (".CmdSrc.Cfg_HasOper", "BOOL", 1),
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
        snapshot_dir=str(tmp_path / "snapshots"),
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


def _upload(client, filename="P_TestValve.L5X", content=None):
    return client.post(
        "/api/aois/import",
        files={"file": (filename, content if content is not None else _L5X.encode())},
    )


def _confirm(client, proposal_id, selected=None):
    body = {"proposal_id": proposal_id}
    if selected is not None:
        body["selected"] = selected
    return client.post("/api/aois/import/confirm", json=body)


def test_upload_returns_proposal(client):
    response = _upload(client)
    assert response.status_code == 200
    proposal = response.json()
    assert proposal["aoi"] == "P_TestValve"
    assert proposal["revision"] == "1.02.00 Test"
    assert proposal["source_file"] == "P_TestValve.L5X"
    assert proposal["nested_objects"] == ["MyAlarm", "CmdSrc"]
    assert proposal["proposal_id"]
    columns = proposal["columns"]
    assert [column["subtag"] for column in columns] == EXPECTED_SUBTAGS
    assert all(column["selected"] is True for column in columns)
    by_subtag = {column["subtag"]: column for column in columns}
    # Pass-through dropped in favour of the nested column; non-Cfg excluded.
    assert ".Cfg_MyAlarmSeverity" not in by_subtag
    assert ".Inp_OpenLS" not in by_subtag
    assert ".Sts_Opened" not in by_subtag
    assert by_subtag[".Cfg_Total"]["type"] == "REAL"
    assert by_subtag[".Cfg_Total"]["default"] == pytest.approx(5.0)
    assert by_subtag[".Cfg_CoastToLS"]["default"] is True
    assert by_subtag[".Cfg_Count"]["default"] == 42
    assert by_subtag[".Cfg_Tag"]["default"] == "P_TestValve"
    assert by_subtag[".MyAlarm.Cfg_Severity"]["type"] == "DINT"


def test_upload_persists_to_import_dir(client, tmp_path):
    assert _upload(client).status_code == 200
    saved = tmp_path / "imports" / "P_TestValve.L5X"
    assert saved.exists()
    assert saved.read_bytes() == _L5X.encode()


def test_upload_rejects_wrong_extension(client):
    response = _upload(client, filename="P_TestValve.txt")
    assert response.status_code == 400
    assert ".L5X" in response.json()["detail"]


def test_upload_rejects_non_xml(client):
    response = _upload(client, content=b"this is not xml at all")
    assert response.status_code == 400


def test_upload_rejects_xml_without_aoi(client):
    response = _upload(client, content=b'<?xml version="1.0"?><root/>')
    assert response.status_code == 400


def test_upload_rejects_empty_body(client):
    assert client.post("/api/aois/import").status_code == 400


def test_upload_rejects_oversize(client):
    big = b"A" * (5 * 1024 * 1024 + 1)
    response = _upload(client, content=big)
    assert response.status_code == 400
    assert "5 MB" in response.json()["detail"]


def test_scan_proposes_skips_and_keeps_id_stable(client, tmp_path):
    imports = tmp_path / "imports"
    imports.mkdir()
    (imports / "P_TestValve.L5X").write_text(_L5X, encoding="utf-8")
    (imports / "broken.L5X").write_text("not xml", encoding="utf-8")
    first = client.get("/api/aois/scan")
    assert first.status_code == 200
    body = first.json()
    assert [p["aoi"] for p in body["proposals"]] == ["P_TestValve"]
    assert len(body["skipped"]) == 1
    assert body["skipped"][0]["file"] == "broken.L5X"
    assert body["skipped"][0]["reason"].startswith("unparseable:")
    proposal_id = body["proposals"][0]["proposal_id"]
    second = client.get("/api/aois/scan").json()
    assert [p["proposal_id"] for p in second["proposals"]] == [proposal_id]


def test_scan_skips_already_imported(client, tmp_path):
    proposal = _upload(client).json()
    assert _confirm(client, proposal["proposal_id"]).status_code == 200
    body = client.get("/api/aois/scan").json()
    assert body["proposals"] == []
    assert body["skipped"] == [{"file": "P_TestValve.L5X", "reason": "already imported"}]


def test_scan_skips_oversize_file(client, tmp_path):
    """The folder scan enforces the same 5 MB cap as the upload path."""
    imports = tmp_path / "imports"
    imports.mkdir()
    (imports / "Big.L5X").write_bytes(_L5X.encode() + b"<!--" + b"x" * (6 * 1024 * 1024) + b"-->")
    body = client.get("/api/aois/scan").json()
    assert body["proposals"] == []
    assert len(body["skipped"]) == 1
    assert body["skipped"][0]["file"] == "Big.L5X"
    assert "exceeds the 5 MB import cap" in body["skipped"][0]["reason"]


def test_scan_needs_no_connection(client, tmp_path):
    imports = tmp_path / "imports"
    imports.mkdir()
    (imports / "P_TestValve.L5X").write_text(_L5X, encoding="utf-8")
    assert client.get("/api/aois/scan").status_code == 200


def test_scan_filename_with_newlines_cannot_inject_aoi(client, tmp_path):
    """A scan-dir file whose name contains newlines must not smuggle an AOI.

    The file name is interpolated into the YAML header comment of the
    confirmed custom AOI; without sanitising, the injected lines would
    define a second AOI (here ``P_Evil``) that bypasses name validation
    and collision checks.
    """
    from pypax.custom_aoi import load_custom_aois

    imports = tmp_path / "imports"
    imports.mkdir()
    evil_name = "A\nP_Evil:\n  - '.Cfg_X'\n#Z.L5X"
    (imports / evil_name).write_text(_L5X, encoding="utf-8")
    body = client.get("/api/aois/scan").json()
    assert [p["aoi"] for p in body["proposals"]] == ["P_TestValve"]
    assert _confirm(client, body["proposals"][0]["proposal_id"]).status_code == 200
    assert load_custom_aois(tmp_path / "custom") == {"P_TestValve": EXPECTED_SUBTAGS}


def test_confirm_creates_custom_aoi(client, tmp_path):
    proposal = _upload(client).json()
    response = _confirm(client, proposal["proposal_id"])
    assert response.status_code == 200
    assert response.json() == {"aoi": "P_TestValve", "columns": len(EXPECTED_SUBTAGS)}
    yaml_path = tmp_path / "custom" / "P_TestValve.yaml"
    assert yaml_path.exists()
    text = yaml_path.read_text(encoding="utf-8")
    for subtag in EXPECTED_SUBTAGS:
        assert f"  - '{subtag}'" in text
    # The proposal is consumed: confirming the same id again is unknown.
    assert _confirm(client, proposal["proposal_id"]).status_code == 404


def test_confirm_selected_filters_columns(client, tmp_path):
    proposal = _upload(client).json()
    wanted = [".Cfg_Desc", ".Cfg_Total", ".MyAlarm.Cfg_Severity"]
    response = _confirm(client, proposal["proposal_id"], selected=wanted)
    assert response.json() == {"aoi": "P_TestValve", "columns": 3}
    text = (tmp_path / "custom" / "P_TestValve.yaml").read_text(encoding="utf-8")
    assert "  - '.Cfg_Desc'" in text
    assert "  - '.Cfg_Total'" in text
    assert "  - '.Cfg_HasStop'" not in text


def test_confirm_rejects_unknown_proposal(client):
    assert _confirm(client, "no-such-id").status_code == 404


def test_confirm_rejects_unknown_selected_subtag(client):
    proposal = _upload(client).json()
    response = _confirm(client, proposal["proposal_id"], selected=[".Cfg_Nope"])
    assert response.status_code == 400


def test_confirm_rejects_empty_selection(client):
    proposal = _upload(client).json()
    assert _confirm(client, proposal["proposal_id"], selected=[]).status_code == 400


def test_confirm_template_collision_is_409(client):
    l5x_p_ain = _L5X.replace("P_TestValve", "P_AIn")
    proposal = _upload(client, filename="P_AIn.L5X", content=l5x_p_ain.encode()).json()
    assert proposal["aoi"] == "P_AIn"
    response = _confirm(client, proposal["proposal_id"])
    assert response.status_code == 409
    assert "P_AIn" in response.json()["detail"]


def test_confirm_reimport_updates_in_place(client):
    """A re-proposal of an already-imported name is an update, not a 409.

    Phase 9c: re-importing rewrites the custom AOI in place. The
    second upload of the unchanged file returns update_available with
    an empty diff, and confirming it succeeds.
    """
    first = _upload(client).json()
    assert _confirm(client, first["proposal_id"]).status_code == 200
    second = _upload(client).json()
    assert second["status"] == "update_available"
    assert second["diff"] == {
        "added": [],
        "removed": [],
        "changed": [],
        "unchanged": 13,
        "before_details_known": True,
    }
    response = _confirm(client, second["proposal_id"])
    assert response.status_code == 200
    assert response.json() == {"aoi": "P_TestValve", "columns": 13, "updated": True}


def test_list_and_read_custom_aoi(connected, tmp_path):
    proposal = _upload(connected).json()
    assert _confirm(connected, proposal["proposal_id"]).status_code == 200
    aois = connected.get("/api/aois").json()
    assert {
        "name": "P_TestValve",
        "instance_count": 1,
        "subtag_count": len(EXPECTED_SUBTAGS),
        "custom": True,
    } in aois
    assert {"name": "P_AIn", "instance_count": 2, "subtag_count": 3, "custom": False} in aois
    body = connected.get("/api/aois/P_TestValve").json()
    assert body["name"] == "P_TestValve"
    assert [column["subtag"] for column in body["columns"]] == EXPECTED_SUBTAGS
    by_subtag = {column["subtag"]: column for column in body["columns"]}
    # Columns are enriched with the proposal's types/descriptions.
    assert by_subtag[".Cfg_Total"]["data_type"] == "REAL"
    assert by_subtag[".MyAlarm.Cfg_Severity"]["data_type"] == "DINT"
    assert by_subtag[".Cfg_Desc"]["data_type"] == "STRING_40"
    assert body["rows"] == [
        {
            "name": "VLV_101",
            "values": {subtag: value for (subtag, _type, value) in VALVE_SEED},
        }
    ]


def test_custom_aoi_paste_and_write_flow(connected):
    proposal = _upload(connected).json()
    assert _confirm(connected, proposal["proposal_id"]).status_code == 200
    response = connected.post(
        "/api/aois/P_TestValve/paste",
        json={"tsv": "7.5\n", "anchor_instance": "VLV_101", "anchor_subtag": ".Cfg_Total"},
    )
    assert response.status_code == 200
    assert response.json()["applied"] == 1
    assert response.json()["rejected"] == []
    write = connected.post("/api/write", json={"plc_name": PLC_NAME})
    assert write.status_code == 200
    assert write.json()["summary"] == {"verified": 1, "mismatch": 0, "failed": 0}
    rows = connected.get("/api/aois/P_TestValve").json()["rows"]
    assert rows[0]["values"][".Cfg_Total"] == pytest.approx(7.5)


def test_delete_custom_aoi(connected):
    proposal = _upload(connected).json()
    assert _confirm(connected, proposal["proposal_id"]).status_code == 200
    response = connected.delete("/api/aois/custom/P_TestValve")
    assert response.status_code == 200
    assert response.json() == {"removed": "P_TestValve"}
    names = [aoi["name"] for aoi in connected.get("/api/aois").json()]
    assert "P_TestValve" not in names
    assert connected.delete("/api/aois/custom/P_TestValve").status_code == 404
    assert connected.delete("/api/aois/custom/P_AIn").status_code == 404


def test_paste_into_untyped_custom_column_is_rejected(client, tmp_path, workbook):
    """Losing the source .L5X untypes the columns; pastes must be rejected.

    Without a type the backend cannot validate a value, and pycomm3
    encodes any truthy value as BOOL true — a garbage paste would
    otherwise write a wrong value to the PLC (caught only afterwards
    as a read-back mismatch). Re-importing the .L5X restores the types.
    """
    proposal = _upload(client).json()
    assert _confirm(client, proposal["proposal_id"]).status_code == 200
    # The source export leaves the import folder (e.g. tidied away).
    (tmp_path / "imports" / "P_TestValve.L5X").unlink()
    # Restart the app against the same dirs: the custom AOI survives
    # (its YAML is in config/aoi/custom) but its columns are untyped.
    restarted = create_app(
        plc_factory=_plc_factory,
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
        snapshot_dir=str(tmp_path / "snapshots"),
    )
    restarted_client = TestClient(restarted)
    assert (
        restarted_client.post("/api/connect", json={"commpath": "10.0.0.1/0"}).json()["connected"]
        is True
    )
    columns = restarted_client.get("/api/aois/P_TestValve").json()["columns"]
    assert all(column["data_type"] == "" for column in columns)
    body = restarted_client.post(
        "/api/aois/P_TestValve/paste",
        json={
            "tsv": "not-a-bool",
            "anchor_instance": "VLV_101",
            "anchor_subtag": ".Cfg_HasStop",
        },
    ).json()
    assert body["applied"] == 0
    assert len(body["rejected"]) == 1
    assert "column type is unknown" in body["rejected"][0]["reason"]
    assert restarted_client.get("/api/changes").json() == []
    assert restarted_client.post("/api/write", json={"plc_name": PLC_NAME}).json() == {
        "results": [],
        "summary": {"verified": 0, "mismatch": 0, "failed": 0},
    }


def test_import_and_remove_are_audited(client):
    """Import confirm and custom-AOI removal leave audit entries.

    The audit log is the record of what changed in the plant
    configuration surface — an import is a config change even though
    it touches no PLC tag.
    """
    proposal = _upload(client).json()
    assert _confirm(client, proposal["proposal_id"]).status_code == 200
    assert client.delete("/api/aois/custom/P_TestValve").status_code == 200
    actions = [e["action"] for e in client.get("/api/audit").json()]
    assert "import" in actions
    assert "remove_custom_aoi" in actions
    entries = {e["action"]: e for e in client.get("/api/audit").json()}
    assert entries["import"]["results"] == {"aoi": "P_TestValve", "columns": len(EXPECTED_SUBTAGS)}
    assert entries["remove_custom_aoi"]["results"] == {"aoi": "P_TestValve"}


def test_startup_scan_proposes(workbook, tmp_path):
    imports = tmp_path / "imports"
    imports.mkdir()
    (imports / "P_TestValve.L5X").write_text(_L5X, encoding="utf-8")
    # Build the app on top of the same dirs: startup scan must propose.
    fresh = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(imports),
        custom_aoi_dir=str(tmp_path / "custom"),
        snapshot_dir=str(tmp_path / "snapshots"),
    )
    state = fresh.state.pypax
    assert len(state.proposals) == 1
    proposal = next(iter(state.proposals.values()))
    assert proposal.aoi == "P_TestValve"
    scanned = TestClient(fresh).get("/api/aois/scan").json()
    assert [p["proposal_id"] for p in scanned["proposals"]] == [proposal.proposal_id]


def test_aoi_import_dir_setting(tmp_path, monkeypatch):
    assert AppSettings().aoi_import_dir == "aoi_imports"
    monkeypatch.setenv("PYPAX_AOI_IMPORT_DIR", str(tmp_path / "elsewhere"))
    assert AppSettings().aoi_import_dir == str(tmp_path / "elsewhere")
