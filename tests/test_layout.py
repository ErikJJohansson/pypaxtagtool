"""Tests for layout constants and AOI sheet detection."""

from pypax.layout import (
    BOTTOM_TAG_ROW,
    DATA_TYPE_ROW,
    NAME_COL,
    SETUP_NUM_INSTANCES_COL,
    SETUP_NUM_SUBTAGS_COL,
    SETUP_TAB_NAME_COL,
    START_COL,
    START_ROW,
    TOP_TAG_ROW,
    get_aoi_list,
)


class FakeBook:
    def __init__(self, sheetnames):
        self.sheetnames = sheetnames


def test_layout_constants_match_template():
    assert (START_ROW, START_COL, NAME_COL) == (10, 5, 3)
    assert (DATA_TYPE_ROW, TOP_TAG_ROW, BOTTOM_TAG_ROW) == (6, 7, 8)
    assert (SETUP_NUM_INSTANCES_COL, SETUP_NUM_SUBTAGS_COL, SETUP_TAB_NAME_COL) == (4, 5, 8)


def test_get_aoi_list_finds_underscore_sheets():
    book = FakeBook(["Setup", "P_AIn", "P_Motor", "L_CPU", "TAG_DESC_EM", "MISC"])
    assert get_aoi_list(book) == ["P_AIn", "P_Motor", "L_CPU"]


def test_get_aoi_list_includes_sif_sheets():
    book = FakeBook(["Setup", "SIF_Recipe", "P_AIn"])
    assert get_aoi_list(book) == ["SIF_Recipe", "P_AIn"]


def test_get_aoi_list_skips_short_and_malformed_names():
    # The old implementation indexed sheet[1]/sheet[2] directly and raised
    # IndexError on 1-character sheet names.
    book = FakeBook(["", "X", "S", "SI", "P_AIn"])
    assert get_aoi_list(book) == ["P_AIn"]


def test_get_aoi_list_empty_workbook():
    assert get_aoi_list(FakeBook([])) == []
