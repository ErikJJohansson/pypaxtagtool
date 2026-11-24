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


def test_coerce_string_keeps_text():
    assert coerce_value("  hello world  ", "STRING_40") == ("  hello world  ", None)


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
    assert changes[0]["subtag"] == ".Cfg_Desc"
    assert changes[1]["subtag"] == ".HiHi.Cfg_Exists"
