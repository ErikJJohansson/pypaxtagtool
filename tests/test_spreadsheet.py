"""Tests for pypax.spreadsheet."""

import openpyxl
import pytest

from helpers import build_workbook
from pypax.spreadsheet import (
    get_aoi_setup,
    get_subtag,
    get_subtag_list,
    read_data_sheet_row,
    search_value_in_col,
    set_num_instances,
    write_sheet_row,
)
from helpers import SUBTAG_NAMES as HELPER_SUBTAGS


@pytest.fixture
def book(tmp_path):
    path = tmp_path / "template.xlsx"
    build_workbook(path)
    workbook = openpyxl.load_workbook(path)
    yield workbook
    workbook.close()


@pytest.fixture
def sheet(book):
    return book["P_AIn"]


def test_get_subtag_list(sheet):
    assert get_subtag_list(sheet) == HELPER_SUBTAGS


def test_get_subtag(sheet):
    assert get_subtag(sheet, 5) == ".Cfg_Desc"
    assert get_subtag(sheet, 7) == ".HiHi.Cfg_Exists"


def test_search_value_in_col(book):
    assert search_value_in_col(book["Setup"], "P_AIn", 8) == 2
    assert search_value_in_col(book["Setup"], "P_Motor", 8) is None


def test_get_aoi_setup_empty_and_filled(tmp_path):
    build_workbook(tmp_path / "blank.xlsx")
    book = openpyxl.load_workbook(tmp_path / "blank.xlsx")
    assert get_aoi_setup(book["P_AIn"]) == 0
    book.close()

    build_workbook(
        tmp_path / "filled.xlsx",
        [("FIT_101", ["Flow", 90.0, 1]), ("PIT_201", ["Pressure", 8.0, 0])],
    )
    book = openpyxl.load_workbook(tmp_path / "filled.xlsx")
    assert get_aoi_setup(book["P_AIn"]) == 2
    book.close()


def test_get_aoi_setup_stops_at_first_blank_row(tmp_path):
    build_workbook(tmp_path / "gap.xlsx")
    book = openpyxl.load_workbook(tmp_path / "gap.xlsx")
    sheet = book["P_AIn"]
    sheet.cell(10, 3).value = "FIT_101"
    sheet.cell(11, 3).value = "PIT_201"
    # Row 12 left blank, row 13 has a name anyway: counting stops at 12.
    sheet.cell(13, 3).value = "LIT_301"
    assert get_aoi_setup(sheet) == 2
    book.close()


def test_set_num_instances(book):
    set_num_instances(book["Setup"], "P_AIn", 7)
    assert book["Setup"].cell(2, 4).value == 7
    set_num_instances(book["Setup"], "P_Motor", 3)  # unknown AOI: no-op
    assert book["Setup"].cell(2, 4).value == 7


def test_write_then_read_sheet_row_round_trip(sheet):
    tag_data = [
        ("FIT_101.Cfg_Desc", "Flow Transmitter 101"),
        ("FIT_101.Cfg_HiLim", 90.0),
        ("FIT_101.HiHi.Cfg_Exists", 1),
    ]
    write_sheet_row(sheet, 10, "FIT_101", tag_data)
    base_tag, read_back = read_data_sheet_row(sheet, 10, HELPER_SUBTAGS)
    assert base_tag == "FIT_101"
    assert read_back == tag_data


def test_read_data_sheet_row_coerces_by_data_type(sheet):
    sheet.cell(10, 3).value = "FIT_101"
    sheet.cell(10, 5).value = "Flow"  # STRING_40: kept as-is
    sheet.cell(10, 6).value = 90  # REAL: int cell -> float
    sheet.cell(10, 7).value = True  # BOOL: -> int 1
    _, tag_data = read_data_sheet_row(sheet, 10, HELPER_SUBTAGS)
    assert tag_data == [
        ("FIT_101.Cfg_Desc", "Flow"),
        ("FIT_101.Cfg_HiLim", 90.0),
        ("FIT_101.HiHi.Cfg_Exists", 1),
    ]
    assert isinstance(tag_data[1][1], float)
    assert isinstance(tag_data[2][1], int)


def test_read_data_sheet_row_blank_cell_becomes_empty_string(sheet):
    sheet.cell(10, 3).value = "FIT_101"
    sheet.cell(10, 5).value = None
    sheet.cell(10, 6).value = 90.0
    sheet.cell(10, 7).value = 0
    _, tag_data = read_data_sheet_row(sheet, 10, HELPER_SUBTAGS)
    assert tag_data[0] == ("FIT_101.Cfg_Desc", "")


def test_read_data_sheet_row_non_string_value_kept(sheet):
    # Regression case from main history: non-string values must survive
    # the write path (the May 2025 "Fixed non string writes" fix).
    sheet.cell(10, 3).value = "FIT_101"
    sheet.cell(10, 5).value = 12345  # a number in a STRING cell: kept as-is
    sheet.cell(10, 6).value = 90.0
    sheet.cell(10, 7).value = 1
    _, tag_data = read_data_sheet_row(sheet, 10, HELPER_SUBTAGS)
    assert tag_data[0] == ("FIT_101.Cfg_Desc", 12345)


def test_functions_work_on_read_only_workbook(tmp_path):
    # Write mode opens files read-only; the row-iteration scans must work there.
    build_workbook(tmp_path / "ro.xlsx", [("FIT_101", ["Flow", 90.0, 1])])
    book = openpyxl.load_workbook(tmp_path / "ro.xlsx", data_only=True, read_only=True)
    sheet = book["P_AIn"]
    assert get_subtag_list(sheet) == HELPER_SUBTAGS
    assert get_aoi_setup(sheet) == 1
    assert search_value_in_col(book["Setup"], "P_AIn", 8) == 2
    base_tag, tag_data = read_data_sheet_row(sheet, 10, HELPER_SUBTAGS)
    assert base_tag == "FIT_101"
    assert tag_data[1] == ("FIT_101.Cfg_HiLim", 90.0)
    book.close()
