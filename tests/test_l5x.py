"""Tests for the Studio 5000 .L5X parser and column curation.

The fixture is a small synthetic export built in-test — never a real
Rockwell file. The real-export test at the bottom reads the P_ValveMO
export Erik supplied (read-only, outside the repo) and checks the
curated proposal against the template sheet.
"""

from pathlib import Path

import openpyxl
import pytest

from pypax.l5x import curate_columns, parse_l5x
from pypax.spreadsheet import get_subtag_list

REPO_ROOT = Path(__file__).resolve().parent.parent
REAL_L5X = Path("/home/hatch/workspace/user/files/P_ValveMO_4.10.06_AOI.L5X")
TEMPLATE = REPO_ROOT / "ProcessLibraryOnlineConfigTool.xlsm"

_FIXTURE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<Parameter Name="Cfg_MyAlarmAckReqd" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Pass-through: use MyAlarm.Cfg_AckReqd]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_MyAlarmResetReqd" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Pass-through: use MyAlarm.Cfg_ResetReqd]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_HasMyAlarmAlm" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Pass-through: use MyAlarm.Cfg_Exists]]></Description>
<DefaultData>00</DefaultData>
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
<Parameter Name="PCmd_Open" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Operator open command]]></Description>
<DefaultData>00</DefaultData>
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
<DefaultData>0A 00 00 00 54 65 73 74 20 56 61 6C 76 65 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00</DefaultData>
</LocalTag>
<LocalTag Name="Cfg_Area" DataType="STRING_Area" ExternalAccess="Read/Write">
<Description><![CDATA[Process Area for security]]></Description>
<DefaultData>06 00 00 00 61 72 65 61 30 31 00 00</DefaultData>
</LocalTag>
<LocalTag Name="Wrk_Foo" DataType="BOOL" ExternalAccess="Read/Write">
<Description><![CDATA[Working tag]]></Description>
<DefaultData>00</DefaultData>
</LocalTag>
<LocalTag Name="Ghost" DataType="P_Unknown" ExternalAccess="Read/Write">
<Description><![CDATA[Nested type not defined in this file]]></Description>
</LocalTag>
</LocalTags>
</AddOnInstructionDefinition>
<AddOnInstructionDefinition Name="P_Alarm" Revision="1.00" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Cfg_Exists" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Alarm exists]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Cfg_ResetReqd" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Reset required]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_AckReqd" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Acknowledge required]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Cfg_Severity" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Alarm severity]]></Description>
<DefaultData>EE 02 00 00</DefaultData>
</Parameter>
<Parameter Name="Cfg_AllowShelve" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[1=Allow shelve]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Cfg_AlmMinOnT" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Minimum on time]]></Description>
<DefaultData>0A 00 00 00</DefaultData>
</Parameter>
<Parameter Name="Cfg_MaxShelfT" TagType="Base" DataType="DINT" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Max shelf time]]></Description>
<DefaultData>00 00 00 00</DefaultData>
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
<LocalTag Name="Wrk_Alm" DataType="BOOL" ExternalAccess="Read/Write">
<Description><![CDATA[Working tag]]></Description>
<DefaultData>00</DefaultData>
</LocalTag>
</LocalTags>
</AddOnInstructionDefinition>
<AddOnInstructionDefinition Name="P_CmdSrc" Revision="1.00" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Cfg_HasOper" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Has operator]]></Description>
<DefaultData>01</DefaultData>
</Parameter>
<Parameter Name="Cfg_HasProg" TagType="Base" DataType="BOOL" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Has program]]></Description>
<DefaultData>00</DefaultData>
</Parameter>
<Parameter Name="Sts_Oper" TagType="Base" DataType="BOOL" Usage="Output" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Operator status]]></Description>
</Parameter>
</Parameters>
</AddOnInstructionDefinition>
</Controller>
</RSLogix5000Content>
"""


@pytest.fixture()
def fixture_path(tmp_path):
    path = tmp_path / "P_TestValve.L5X"
    path.write_text(_FIXTURE, encoding="utf-8")
    return path


def test_parse_synthetic(fixture_path):
    definition = parse_l5x(fixture_path)
    assert definition.name == "P_TestValve"
    assert definition.revision == "1.02.00 Test"
    assert set(definition.sub_aois) == {"P_Alarm", "P_CmdSrc"}
    members = {member.name: member for member in definition.members}
    assert members["Cfg_CoastToLS"].default is True
    assert members["Cfg_HasStop"].default is False
    assert members["Cfg_Total"].default == pytest.approx(5.0)
    assert members["Cfg_Total"].data_type == "REAL"
    assert members["Cfg_Count"].default == 42
    assert members["Cfg_Desc"].default == "Test Valve"
    assert members["Cfg_Label"].default == "Test Valve"
    assert members["Cfg_Tag"].default == "P_TestValve"
    assert members["Cfg_Area"].default == "area01"
    assert members["Cfg_CoastToLS"].description == "1=Coasting into LS when stopped"
    # No DefaultData element at all -> None, not a crash.
    assert members["Sts_Opened"].default is None
    alarm = definition.sub_aois["P_Alarm"]
    alarm_members = {member.name: member for member in alarm.members}
    assert alarm_members["Cfg_Severity"].default == 750
    assert alarm_members["Cfg_Cond"].default == "Alarm Condition"
    assert alarm_members["Cfg_Tag"].default == "P_Alarm"


def test_curate_columns_exact_order(fixture_path):
    columns = curate_columns(parse_l5x(fixture_path))
    assert [column.subtag for column in columns] == [
        ".Cfg_Desc",
        ".Cfg_Label",
        ".Cfg_Tag",
        ".Cfg_Area",
        ".Cfg_HasStop",
        ".Cfg_CoastToLS",
        ".Cfg_Total",
        ".Cfg_Count",
        ".MyAlarm.Cfg_Exists",
        ".MyAlarm.Cfg_ResetReqd",
        ".MyAlarm.Cfg_AckReqd",
        ".MyAlarm.Cfg_Severity",
        ".MyAlarm.Cfg_AllowShelve",
        ".MyAlarm.Cfg_AlmMinOnT",
        ".MyAlarm.Cfg_MaxShelfT",
        ".MyAlarm.Cfg_Tag",
        ".MyAlarm.Cfg_Cond",
        ".CmdSrc.Cfg_HasOper",
        ".CmdSrc.Cfg_HasProg",
    ]


def test_curation_drops_passthroughs_keeps_nested(fixture_path):
    subtags = {column.subtag for column in curate_columns(parse_l5x(fixture_path))}
    for dropped in (
        ".Cfg_MyAlarmSeverity",
        ".Cfg_MyAlarmAckReqd",
        ".Cfg_MyAlarmResetReqd",
        ".Cfg_HasMyAlarmAlm",
    ):
        assert dropped not in subtags
    for kept in (
        ".MyAlarm.Cfg_Severity",
        ".MyAlarm.Cfg_AckReqd",
        ".MyAlarm.Cfg_ResetReqd",
        ".MyAlarm.Cfg_Exists",
    ):
        assert kept in subtags


def test_curation_excludes_non_config(fixture_path):
    subtags = {column.subtag for column in curate_columns(parse_l5x(fixture_path))}
    for excluded in (".Inp_OpenLS", ".Sts_Opened", ".PCmd_Open", ".Wrk_Foo"):
        assert excluded not in subtags
    # A nested head whose DataType is not defined in the file expands
    # to nothing and never crashes the import.
    assert not any(subtag.startswith(".Ghost") for subtag in subtags)
    assert all(subtag.lstrip(".").split(".")[-1].startswith("Cfg_") for subtag in subtags)


def test_curation_column_carries_type_description_default(fixture_path):
    by_subtag = {column.subtag: column for column in curate_columns(parse_l5x(fixture_path))}
    assert by_subtag[".Cfg_Total"].data_type == "REAL"
    assert by_subtag[".Cfg_Total"].default == pytest.approx(5.0)
    assert by_subtag[".Cfg_Count"].default == 42
    assert by_subtag[".MyAlarm.Cfg_Severity"].default == 750
    assert by_subtag[".Cfg_CoastToLS"].description == "1=Coasting into LS when stopped"


def test_parse_rejects_bad_files(tmp_path):
    bad = tmp_path / "bad.L5X"
    bad.write_text("this is not xml", encoding="utf-8")
    with pytest.raises(ValueError, match="not a readable"):
        parse_l5x(bad)
    empty = tmp_path / "empty.L5X"
    empty.write_text("<RSLogix5000Content></RSLogix5000Content>", encoding="utf-8")
    with pytest.raises(ValueError, match="no AddOnInstructionDefinition"):
        parse_l5x(empty)
    with pytest.raises(ValueError, match="cannot read"):
        parse_l5x(tmp_path / "missing.L5X")


def test_real_valve_export_proposal_matches_sheet_plus_intlktrip():
    """The curated proposal for the real P_ValveMO export.

    The template sheet's 82 subtags must all be proposed, and the only
    extras are the IntlkTrip alarm columns (the sheet omits that alarm
    by choice) plus ``.Cfg_HasIntlkObj`` — a genuine top-level ``Cfg_``
    parameter the sheet also omits, which is not a pass-through of any
    nested object, so the curation rule keeps it. That last one is a
    spec ambiguity flagged for arbitration, not a code bug: the sheet
    includes the other ``Cfg_Has*Obj`` flags, so no mechanical rule
    reproduces its omission of this one.
    """
    if not REAL_L5X.exists():
        pytest.skip(f"real export not present: {REAL_L5X}")
    definition = parse_l5x(REAL_L5X)
    assert definition.name == "P_ValveMO"
    assert "4.10" in definition.revision

    proposal = {column.subtag for column in curate_columns(definition)}

    book = openpyxl.load_workbook(TEMPLATE, read_only=True, data_only=True)
    try:
        sheet = set(get_subtag_list(book["P_ValveMO"]))
    finally:
        book.close()
    assert len(sheet) == 82
    assert sheet <= proposal, f"sheet subtags missing from proposal: {sheet - proposal}"

    alarm_leaves = {
        member.name
        for member in definition.sub_aois["P_Alarm"].members
        if member.name.startswith("Cfg_")
    }
    expected_extras = {f".IntlkTrip.{leaf}" for leaf in alarm_leaves}
    expected_extras.add(".Cfg_HasIntlkObj")
    assert proposal - sheet == expected_extras, (
        f"unexpected proposal extras: {(proposal - sheet) - expected_extras}; "
        f"missing expected extras: {expected_extras - (proposal - sheet)}"
    )
    # The proposal is a Column list, not just names: spot-check metadata.
    by_subtag = {column.subtag: column for column in curate_columns(definition)}
    assert by_subtag[".Cfg_CoastToLS"].data_type == "BOOL"
    assert by_subtag[".Cfg_CoastToLS"].default is False
    assert by_subtag[".Cfg_FullStallT"].default == pytest.approx(15.0)
    assert by_subtag[".Cfg_Desc"].default == "Motor Operated Valve"
