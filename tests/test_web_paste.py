"""Unit tests for TSV paste parsing and value coercion."""

from __future__ import annotations

from pypax.web.aoi_model import Column
from pypax.web.paste import coerce_value, parse_tsv

COLUMNS = [
    Column(".Cfg_Desc", "STRING_40", ""),
    Column(".Cfg_HiLim", "REAL", ""),
    Column(".HiHi.Cfg_Exists", "BOOL", ""),
]
INSTANCES = ["FIT_101", "PIT_201"]


def test_coerce_bool_variants():
    for raw, expected in [
        ("1", 1),
        ("0", 0),
        ("true", 1),
        ("FALSE", 0),
        ("Yes", 1),
        ("off", 0),
    ]:
        assert coerce_value(raw, "BOOL") == (expected, None)


def test_coerce_bool_rejected():
    value, reason = coerce_value("maybe", "BOOL")
    assert value is None
    assert "not a valid BOOL" in reason


def test_coerce_int_rejects_float_text():
    assert coerce_value("42", "DINT") == (42, None)
    assert coerce_value("-7", "INT") == (-7, None)
    value, reason = coerce_value("4.5", "DINT")
    assert value is None and "not a valid DINT" in reason


def test_coerce_real_rejects_nan_and_text():
    assert coerce_value("3.5", "REAL") == (3.5, None)
    assert coerce_value("90", "REAL") == (90.0, None)
    for bad in ("abc", "nan", "inf"):
        value, reason = coerce_value(bad, "REAL")
        assert value is None and "not a valid REAL" in reason


def test_coerce_rejects_unknown_type():
    """An untyped column accepts nothing.

    Only a custom AOI whose source .L5X is gone has no type, and
    without a type the value cannot be validated — pycomm3 would
    encode any truthy value (including garbage) as BOOL true.
    """
    value, reason = coerce_value("1", "")
    assert value is None
    assert "column type is unknown" in reason


def test_parse_rejects_untyped_columns():
    columns = [Column(".Cfg_HasStop", "", "")]
    changes, rejected = parse_tsv("1", INSTANCES, columns)
    assert changes == []
    assert len(rejected) == 1
    assert "column type is unknown" in rejected[0]["reason"]


def test_coerce_string_keeps_text():
    assert coerce_value("  hello world  ", "STRING_40") == ("  hello world  ", None)


def test_coerce_string_learned_size_from_registry(tmp_path):
    """A size learned from an .L5X import (tmp registry file) is enforced."""
    registry = tmp_path / "string_udt_sizes.yaml"
    registry.write_text(
        "STRING_FOO:\n  size: 30\n  source: P_Foo_1.0_AOI.L5X\n",
        encoding="utf-8",
    )
    assert coerce_value("x" * 30, "STRING_FOO", registry) == ("x" * 30, None)
    value, reason = coerce_value("x" * 31, "STRING_FOO", registry)
    assert value is None and "30-character limit" in reason


def test_coerce_string_pinned_wins_over_learned(tmp_path):
    """A hand-written learned entry conflicting with a pinned size loses."""
    registry = tmp_path / "string_udt_sizes.yaml"
    registry.write_text(
        "STRING_COND:\n  size: 99\n  source: HandEdited.L5X\n",
        encoding="utf-8",
    )
    assert coerce_value("x" * 20, "STRING_COND", registry) == ("x" * 20, None)
    value, reason = coerce_value("x" * 21, "STRING_COND", registry)
    assert value is None and "20-character limit" in reason


def test_coerce_string_missing_registry_falls_back_to_pinned(tmp_path):
    """No registry file at all: pinned sizes still apply, unknown types unchecked."""
    missing = tmp_path / "no-such-file.yaml"
    assert coerce_value("x" * 8, "STRING_AREA", missing) == ("x" * 8, None)
    value, reason = coerce_value("x" * 9, "STRING_AREA", missing)
    assert value is None and "8-character limit" in reason
    assert coerce_value("x" * 100, "STRING_FOO", missing) == ("x" * 100, None)


def test_coerce_string_length_limits():
    """A string longer than its type holds is rejected, like the
    spreadsheet flags it: STRING_N holds N, the Logix STRING holds 82,
    STRING_Area holds 8 (DATA SINT[8] in the P_ValveMO export),
    STRING_Cond and STRING_NavTag hold 20 (DATA SINT[20] in the
    P_Intlk export)."""
    assert coerce_value("12345678", "STRING_8") == ("12345678", None)
    value, reason = coerce_value("123456789", "STRING_8")
    assert value is None and "9 characters" in reason and "8-character limit" in reason
    assert coerce_value("x" * 40, "STRING_40") == ("x" * 40, None)
    value, reason = coerce_value("x" * 41, "STRING_40")
    assert value is None and "41 characters" in reason and "40-character limit" in reason
    assert coerce_value("x" * 82, "STRING") == ("x" * 82, None)
    value, reason = coerce_value("x" * 83, "STRING")
    assert value is None and "82-character limit" in reason
    assert coerce_value("area01", "STRING_AREA") == ("area01", None)
    value, reason = coerce_value("area01234", "STRING_AREA")
    assert value is None and "8-character limit" in reason
    assert coerce_value("x" * 20, "STRING_COND") == ("x" * 20, None)
    value, reason = coerce_value("x" * 21, "STRING_COND")
    assert value is None and "20-character limit" in reason
    assert coerce_value("x" * 20, "STRING_NAVTAG") == ("x" * 20, None)
    value, reason = coerce_value("x" * 21, "STRING_NAVTAG")
    assert value is None and "20-character limit" in reason


def test_parse_positional_rows():
    changes, rejected = parse_tsv("A desc\t1.5\t1\nB desc\t2.5\t0\n", INSTANCES, COLUMNS)
    assert rejected == []
    assert [c["instance"] for c in changes] == [
        "FIT_101",
        "FIT_101",
        "FIT_101",
        "PIT_201",
        "PIT_201",
        "PIT_201",
    ]
    assert changes[0]["tag"] == "FIT_101.Cfg_Desc"
    assert changes[1]["value"] == 1.5


def test_parse_instance_named_rows():
    changes, rejected = parse_tsv("PIT_201\tOnly this\t9.5\t\n", INSTANCES, COLUMNS)
    assert rejected == []
    assert len(changes) == 2
    assert all(c["instance"] == "PIT_201" for c in changes)


def test_parse_too_many_rows_rejected():
    _changes, rejected = parse_tsv("a\t1\t1\nb\t2\t0\nc\t3\t1\n", INSTANCES, COLUMNS)
    assert rejected == [{"row": "c", "column": "", "reason": "no instance for pasted row 3"}]


def test_parse_anchor_maps_block_start():
    changes, rejected = parse_tsv(
        "9.5\t1\n",
        INSTANCES,
        COLUMNS,
        anchor_instance="PIT_201",
        anchor_subtag=".Cfg_HiLim",
    )
    assert rejected == []
    assert [(c["instance"], c["subtag"], c["value"]) for c in changes] == [
        ("PIT_201", ".Cfg_HiLim", 9.5),
        ("PIT_201", ".HiHi.Cfg_Exists", 1),
    ]


def test_parse_anchor_subtag_only_shifts_columns():
    changes, rejected = parse_tsv("9.5\n", INSTANCES, COLUMNS, anchor_subtag=".Cfg_HiLim")
    assert rejected == []
    assert [(c["instance"], c["subtag"]) for c in changes] == [("FIT_101", ".Cfg_HiLim")]


def test_parse_unknown_anchor_rejected():
    changes, rejected = parse_tsv("1\n", INSTANCES, COLUMNS, anchor_instance="NOPE_999")
    assert changes == []
    assert rejected == [{"row": "NOPE_999", "column": "", "reason": "unknown instance 'NOPE_999'"}]


def test_parse_rejected_rows_use_instance_names():
    _changes, rejected = parse_tsv("Desc\tnope\t1\n", INSTANCES, COLUMNS)
    assert rejected == [
        {
            "row": "FIT_101",
            "column": ".Cfg_HiLim",
            "reason": "not a valid REAL, got 'nope'",
        }
    ]


def test_parse_skips_blank_lines_and_cells():
    changes, rejected = parse_tsv("\nA\t\t1\n\n", INSTANCES, COLUMNS)
    assert rejected == []
    assert len(changes) == 2
    # The leading/trailing blank lines are not part of the block: the
    # values land on the first instance, not shifted down one row.
    assert all(c["instance"] == "FIT_101" for c in changes)
    assert changes[0]["subtag"] == ".Cfg_Desc"
    assert changes[1]["subtag"] == ".HiHi.Cfg_Exists"


def test_parse_blank_line_inside_block_keeps_row_mapping():
    """A blank line in the middle of a paste is an empty row.

    It advances the row position: the values after it must land on the
    instance below the blank row, not silently shift up onto the wrong
    instance.
    """
    instances = ["FIT_101", "PIT_201", "TIT_301"]
    changes, rejected = parse_tsv("A\t11.0\n\nB\t13.0", instances, COLUMNS)
    assert rejected == []
    assert [(c["instance"], c["subtag"], c["value"]) for c in changes] == [
        ("FIT_101", ".Cfg_Desc", "A"),
        ("FIT_101", ".Cfg_HiLim", 11.0),
        ("TIT_301", ".Cfg_Desc", "B"),
        ("TIT_301", ".Cfg_HiLim", 13.0),
    ]


def test_parse_blank_line_inside_block_keeps_anchor_mapping():
    instances = ["FIT_101", "PIT_201", "TIT_301"]
    changes, rejected = parse_tsv(
        "11.0\n\n13.0", instances, COLUMNS, anchor_instance="FIT_101", anchor_subtag=".Cfg_HiLim"
    )
    assert rejected == []
    assert [(c["instance"], c["value"]) for c in changes] == [
        ("FIT_101", 11.0),
        ("TIT_301", 13.0),
    ]
