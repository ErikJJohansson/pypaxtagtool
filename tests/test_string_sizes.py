"""Tests for learning string-UDT sizes from .L5X imports (Phase 9b).

The parser fixture is a small synthetic type table built in-test —
never a real Rockwell file. Registry tests use tmp files; the
committed ``config/string_udt_sizes.yaml`` is only read, never
written, by these tests.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from pypax.l5x import learn_string_sizes
from pypax.string_sizes import (
    PINNED_SIZES,
    load_learned_sizes,
    merge_learned_sizes,
    merged_sizes,
)

_FIXTURE_TYPES = """<?xml version="1.0" encoding="UTF-8"?>
<RSLogix5000Content TargetName="P_Test">
<DataTypes>
<DataType Name="STRING_Foo" Family="StringFamily" Class="User">
<Members>
<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal"/>
<Member Name="DATA" DataType="SINT" Dimension="33" Radix="ASCII"/>
</Members>
</DataType>
<DataType Name="STRING_Bar" Family="StringFamily" Class="User">
<Members>
<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal"/>
<Member Name="DATA" DataType="SINT" Dimension="12" Radix="ASCII"/>
</Members>
</DataType>
<DataType Name="STRING_NoData" Family="StringFamily" Class="User">
<Members>
<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal"/>
</Members>
</DataType>
<DataType Name="STRING_BadData" Family="StringFamily" Class="User">
<Members>
<Member Name="DATA" DataType="DINT" Dimension="40" Radix="ASCII"/>
</Members>
</DataType>
<DataType Name="STRING_NoDim" Family="StringFamily" Class="User">
<Members>
<Member Name="DATA" DataType="SINT" Radix="ASCII"/>
</Members>
</DataType>
<DataType Name="MyUdt" Family="User" Class="User">
<Members>
<Member Name="DATA" DataType="SINT" Dimension="99" Radix="ASCII"/>
</Members>
</DataType>
<DataType Name="STRING_lower" Family="StringFamily" Class="User">
<Members>
<Member Name="DATA" DataType="SINT" Dimension="7" Radix="ASCII"/>
</Members>
</DataType>
</DataTypes>
</RSLogix5000Content>
"""


def test_learn_string_sizes_from_type_table():
    """Only StringFamily UDTs with a SINT DATA member count, keyed upper-cased."""
    root = ET.fromstring(_FIXTURE_TYPES)
    assert learn_string_sizes(root) == {
        "STRING_FOO": 33,
        "STRING_BAR": 12,
        "STRING_LOWER": 7,
    }


def test_learn_string_sizes_empty_when_no_string_types():
    root = ET.fromstring(
        "<RSLogix5000Content><DataTypes>"
        '<DataType Name="MyUdt" Family="User"><Members>'
        '<Member Name="DATA" DataType="SINT" Dimension="9"/>'
        "</Members></DataType></DataTypes></RSLogix5000Content>"
    )
    assert learn_string_sizes(root) == {}


def test_learn_string_sizes_last_definition_wins():
    root = ET.fromstring(
        "<RSLogix5000Content><DataTypes>"
        '<DataType Name="STRING_Dup" Family="StringFamily"><Members>'
        '<Member Name="DATA" DataType="SINT" Dimension="10"/>'
        "</Members></DataType>"
        '<DataType Name="STRING_Dup" Family="StringFamily"><Members>'
        '<Member Name="DATA" DataType="SINT" Dimension="25"/>'
        "</Members></DataType></DataTypes></RSLogix5000Content>"
    )
    assert learn_string_sizes(root) == {"STRING_DUP": 25}


def test_load_learned_sizes_missing_file_is_empty(tmp_path):
    assert load_learned_sizes(tmp_path / "nope.yaml") == {}


def test_load_learned_sizes_malformed_file_is_empty(tmp_path):
    path = tmp_path / "bad.yaml"
    path.write_text("STRING_FOO: [unclosed", encoding="utf-8")
    assert load_learned_sizes(path) == {}
    path.write_text("- just\n- a\n- list\n", encoding="utf-8")
    assert load_learned_sizes(path) == {}


def test_load_learned_sizes_skips_malformed_entries(tmp_path):
    path = tmp_path / "registry.yaml"
    path.write_text(
        "STRING_OK:\n  size: 30\n  source: A.L5X\n"
        "STRING_NO_SIZE:\n  source: A.L5X\n"
        "STRING_ZERO:\n  size: 0\n  source: A.L5X\n"
        "STRING_TEXT_SIZE:\n  size: big\n  source: A.L5X\n"
        "STRING_NO_SOURCE:\n  size: 10\n"
        "bad-name:\n  size: 10\n  source: A.L5X\n",
        encoding="utf-8",
    )
    assert load_learned_sizes(path) == {"STRING_OK": 30}


def test_merge_learned_sizes_persists_with_source(tmp_path):
    path = tmp_path / "registry.yaml"
    warnings = merge_learned_sizes({"STRING_FOO": 30}, "P_Foo_1.0_AOI.L5X", path=path)
    assert warnings == []
    assert load_learned_sizes(path) == {"STRING_FOO": 30}
    text = path.read_text(encoding="utf-8")
    assert "P_Foo_1.0_AOI.L5X" in text
    assert "size: 30" in text


def test_merge_learned_sizes_pinned_conflict_ignored_with_warning(tmp_path):
    path = tmp_path / "registry.yaml"
    warnings = merge_learned_sizes({"STRING_COND": 99}, "P_X_1.0_AOI.L5X", path=path)
    assert len(warnings) == 1 and "STRING_COND" in warnings[0] and "pinned" in warnings[0]
    assert load_learned_sizes(path) == {}
    assert not path.exists()
    # Same size as pinned: no warning, nothing stored.
    assert merge_learned_sizes({"STRING_COND": 20}, "P_X_1.0_AOI.L5X", path=path) == []
    assert not path.exists()


def test_merge_learned_sizes_learned_conflict_keeps_first(tmp_path):
    path = tmp_path / "registry.yaml"
    assert merge_learned_sizes({"STRING_FOO": 30}, "A.L5X", path=path) == []
    warnings = merge_learned_sizes({"STRING_FOO": 40}, "B.L5X", path=path)
    assert len(warnings) == 1 and "STRING_FOO" in warnings[0]
    assert load_learned_sizes(path) == {"STRING_FOO": 30}
    # Same size again: silent, no rewrite.
    assert merge_learned_sizes({"STRING_FOO": 30}, "B.L5X", path=path) == []


def test_merged_sizes_pinned_wins_over_hand_written_conflict(tmp_path):
    path = tmp_path / "registry.yaml"
    path.write_text(
        "STRING_COND:\n  size: 99\n  source: HandEdited.L5X\n"
        "STRING_FOO:\n  size: 30\n  source: A.L5X\n",
        encoding="utf-8",
    )
    merged = merged_sizes(path)
    assert merged["STRING_COND"] == PINNED_SIZES["STRING_COND"] == 20
    assert merged["STRING_FOO"] == 30
    assert merged["STRING"] == 82
    assert merged["STRING_AREA"] == 8


def test_committed_registry_seeds_known_sizes():
    """The repo's own registry carries the two verified sizes."""
    path = Path(__file__).resolve().parent.parent / "config" / "string_udt_sizes.yaml"
    assert load_learned_sizes(path) == {"STRING_COND": 20, "STRING_NAVTAG": 20}


# --- API tests: /api/string-limits + confirm-time learning ------------

_L5X_WITH_FOO = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<RSLogix5000Content SchemaRevision="1.0" TargetName="P_TestValve" TargetType="AddOnInstructionDefinition" TargetRevision="1.00 Test">
<DataTypes>
<DataType Name="STRING_Foo" Family="StringFamily" Class="User">
<Members>
<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal"/>
<Member Name="DATA" DataType="SINT" Dimension="33" Radix="ASCII"/>
</Members>
</DataType>
</DataTypes>
<Controller Use="Context" Name="Test">
<AddOnInstructionDefinition Name="P_TestValve" Revision="1.00" Vendor="Test" ExecutePrescan="true" ExecutePostscan="false" ExecuteEnableInFalse="true">
<Parameters>
<Parameter Name="Cfg_Note" TagType="Base" DataType="STRING_Foo" Usage="Input" Radix="Decimal" Required="false" Visible="false" ExternalAccess="Read/Write">
<Description><![CDATA[Note]]></Description>
</Parameter>
</Parameters>
</AddOnInstructionDefinition>
</Controller>
</RSLogix5000Content>
"""

_L5X_CONFLICT = _L5X_WITH_FOO.replace(
    '<DataType Name="STRING_Foo"',
    '<DataType Name="STRING_Cond"',
).replace('Dimension="33"', 'Dimension="99"')


@pytest.fixture()
def api_app(tmp_path):
    """App with tmp import/custom dirs and a tmp string registry."""
    from fastapi.testclient import TestClient

    from fakes.fake_plc import FakePLC
    from helpers import build_workbook
    from pypax.web import create_app

    template = tmp_path / "template.xlsx"
    build_workbook(template, instance_rows=[("FIT_101", ["Flow", 90.0, 1])])
    registry = tmp_path / "string_udt_sizes.yaml"
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(template),
        audit_dir=str(tmp_path / "audit"),
        import_dir=str(tmp_path / "imports"),
        custom_aoi_dir=str(tmp_path / "custom"),
        string_registry_path=str(registry),
    )
    return TestClient(app), registry


def test_string_limits_endpoint_returns_merged_map(api_app):
    client, _registry = api_app
    response = client.get("/api/string-limits")
    assert response.status_code == 200
    limits = response.json()
    assert limits["STRING"] == 82
    assert limits["STRING_AREA"] == 8
    assert limits["STRING_COND"] == 20
    assert limits["STRING_NAVTAG"] == 20
    # STRING_N is a naming rule, not a map entry.
    assert "STRING_40" not in limits


def test_confirm_learns_string_sizes(api_app):
    client, registry = api_app
    upload = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", _L5X_WITH_FOO.encode())},
    )
    assert upload.status_code == 200
    confirm = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload.json()["proposal_id"]},
    )
    assert confirm.status_code == 200
    assert confirm.json().get("warnings", []) == []
    assert load_learned_sizes(registry) == {"STRING_FOO": 33}
    assert "P_TestValve.L5X" in registry.read_text(encoding="utf-8")
    limits = client.get("/api/string-limits").json()
    assert limits["STRING_FOO"] == 33


def test_reimport_learns_string_sizes(api_app):
    """F1: the re-import (update) path merges learned sizes too — the
    9b merge must not live only on the fresh-import path."""
    client, registry = api_app
    upload = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", _L5X_WITH_FOO.encode())},
    )
    assert upload.status_code == 200
    confirm = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload.json()["proposal_id"]},
    )
    assert confirm.status_code == 200
    assert load_learned_sizes(registry) == {"STRING_FOO": 33}
    # v2: same AOI, a new learned type (STRING_Bar, 7) and a changed
    # description so the diff is non-empty.
    v2 = (
        _L5X_WITH_FOO.replace("STRING_Foo", "STRING_Bar")
        .replace('Dimension="33"', 'Dimension="7"')
        .replace(
            "<Description><![CDATA[Note]]></Description>",
            "<Description><![CDATA[Note v2]]></Description>",
        )
    )
    upload2 = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", v2.encode())},
    )
    assert upload2.status_code == 200
    assert upload2.json()["status"] == "update_available"
    confirm2 = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload2.json()["proposal_id"]},
    )
    assert confirm2.status_code == 200
    assert load_learned_sizes(registry) == {"STRING_FOO": 33, "STRING_BAR": 7}
    assert client.get("/api/string-limits").json()["STRING_BAR"] == 7


def test_reimport_pinned_conflict_warns(api_app):
    """F1: a re-import whose learned size conflicts with a pinned size
    warns in the update response and stores nothing."""
    client, registry = api_app
    upload = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", _L5X_WITH_FOO.encode())},
    )
    assert upload.status_code == 200
    confirm = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload.json()["proposal_id"]},
    )
    assert confirm.status_code == 200
    conflict_v2 = _L5X_CONFLICT.replace(
        "<Description><![CDATA[Note]]></Description>",
        "<Description><![CDATA[Note v2]]></Description>",
    )
    upload2 = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", conflict_v2.encode())},
    )
    assert upload2.status_code == 200
    assert upload2.json()["status"] == "update_available"
    confirm2 = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload2.json()["proposal_id"]},
    )
    assert confirm2.status_code == 200
    warnings = confirm2.json()["warnings"]
    assert len(warnings) == 1
    assert "STRING_COND" in warnings[0] and "pinned" in warnings[0]
    assert load_learned_sizes(registry) == {"STRING_FOO": 33}
    assert client.get("/api/string-limits").json()["STRING_COND"] == 20


def test_confirm_pinned_conflict_warns_and_stores_nothing(api_app):
    client, registry = api_app
    upload = client.post(
        "/api/aois/import",
        files={"file": ("P_TestValve.L5X", _L5X_CONFLICT.encode())},
    )
    assert upload.status_code == 200
    confirm = client.post(
        "/api/aois/import/confirm",
        json={"proposal_id": upload.json()["proposal_id"]},
    )
    assert confirm.status_code == 200
    warnings = confirm.json()["warnings"]
    assert len(warnings) == 1
    assert "STRING_COND" in warnings[0] and "pinned" in warnings[0]
    assert not registry.exists()
    assert client.get("/api/string-limits").json()["STRING_COND"] == 20
