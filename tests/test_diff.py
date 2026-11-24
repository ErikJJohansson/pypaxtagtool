"""Tests for pypax.diff."""

from pypax.diff import row_differences


def test_no_differences_when_rows_match():
    sheet = [("A.Cfg_HiLim", 90), ("A.Cfg_Desc", "Flow")]
    plc = [("A.Cfg_HiLim", 90), ("A.Cfg_Desc", "Flow")]
    assert row_differences(sheet, plc) == []


def test_differences_preserve_spreadsheet_order():
    sheet = [("A.Cfg_HiLim", 42.5), ("A.Cfg_Desc", "New"), ("A.Cfg_LoLim", 5.0)]
    plc = [("A.Cfg_HiLim", 90), ("A.Cfg_Desc", "Old"), ("A.Cfg_LoLim", 5.0)]
    # Sheet order, not PLC order and not set order.
    assert row_differences(sheet, plc) == [
        ("A.Cfg_HiLim", 42.5),
        ("A.Cfg_Desc", "New"),
    ]


def test_duplicate_sheet_tags_kept_once():
    sheet = [("A.Cfg_HiLim", 42.5), ("A.Cfg_HiLim", 42.5)]
    plc = [("A.Cfg_HiLim", 90)]
    assert row_differences(sheet, plc) == [("A.Cfg_HiLim", 42.5)]


def test_sheet_tag_missing_from_plc_is_a_difference():
    sheet = [("A.Cfg_HiLim", 42.5)]
    plc = []
    assert row_differences(sheet, plc) == [("A.Cfg_HiLim", 42.5)]


def test_plc_only_tags_are_ignored():
    sheet = [("A.Cfg_HiLim", 90)]
    plc = [("A.Cfg_HiLim", 90), ("A.Cfg_Other", 1)]
    assert row_differences(sheet, plc) == []


def test_empty_rows():
    assert row_differences([], []) == []
    assert row_differences([], [("A.Cfg_HiLim", 90)]) == []


def test_comparison_is_exact_for_now():
    # 90.0 == 90 in Python, so no diff; a tiny REAL change is a diff.
    # Tolerance handling is Phase 2; Phase 1 keeps exact comparison.
    assert row_differences([("A.Cfg_HiLim", 90.0)], [("A.Cfg_HiLim", 90)]) == []
    assert row_differences([("A.Cfg_HiLim", 42.5)], [("A.Cfg_HiLim", 42.5000001)]) == [
        ("A.Cfg_HiLim", 42.5)
    ]
    # BOOL-as-int compares equal to a real bool.
    assert row_differences([("A.HiHi.Cfg_Exists", 1)], [("A.HiHi.Cfg_Exists", True)]) == []
