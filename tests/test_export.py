"""Unit tests for the spreadsheet export builder."""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from helpers import SUBTAG_NAMES, build_workbook
from pypax.export import AMBER_FILL, build_export_workbook
from pypax.spreadsheet import get_aoi_setup, get_subtag_list, read_data_sheet_row

BASELINES = {
    "P_AIn": {
        "FIT_101": {
            ".Cfg_Desc": "Flow Transmitter 101",
            ".Cfg_HiLim": 90,
            ".HiHi.Cfg_Exists": 1,
        },
        "PIT_201": {
            ".Cfg_Desc": "Pressure Transmitter 201",
            ".Cfg_HiLim": 8,
            ".HiHi.Cfg_Exists": 0,
        },
    }
}

STAGED = [
    {
        "aoi": "P_AIn",
        "instance": "FIT_101",
        "subtag": ".Cfg_HiLim",
        "tag": "FIT_101.Cfg_HiLim",
        "old_value": 90,
        "new_value": 95,
    }
]


def _save_and_reload(book):
    """Round-trip a workbook through bytes, as the endpoint serves it."""
    buf = BytesIO()
    book.save(buf)
    book.close()
    buf.seek(0)
    return openpyxl.load_workbook(buf)


@pytest.fixture()
def template(tmp_path):
    path = tmp_path / "template.xlsx"
    return build_workbook(
        path,
        instance_rows=[
            ("FIT_101", ["Flow Transmitter 101", 90.0, 1]),
            ("PIT_201", ["Pressure Transmitter 201", 8.0, 0]),
        ],
    )


def test_export_writes_baseline_values(template):
    book = _save_and_reload(build_export_workbook(template, BASELINES, [], "TestPLC"))
    sheet = book["P_AIn"]
    assert get_subtag_list(sheet) == SUBTAG_NAMES
    assert get_aoi_setup(sheet) == 2
    base, tag_data = read_data_sheet_row(sheet, 10, SUBTAG_NAMES)
    assert base == "FIT_101"
    assert tag_data == [
        ("FIT_101.Cfg_Desc", "Flow Transmitter 101"),
        ("FIT_101.Cfg_HiLim", 90.0),
        ("FIT_101.HiHi.Cfg_Exists", 1),
    ]
    base, tag_data = read_data_sheet_row(sheet, 11, SUBTAG_NAMES)
    assert base == "PIT_201"
    assert tag_data[1] == ("PIT_201.Cfg_HiLim", 8.0)
    # Setup sheet instance count matches the exported rows.
    assert book["Setup"].cell(2, 4).value == 2
    book.close()


def test_export_applies_staged_values_with_amber_fill(template):
    book = _save_and_reload(build_export_workbook(template, BASELINES, STAGED, "TestPLC"))
    sheet = book["P_AIn"]
    # Staged HiLim (col F) shows the new value with the amber fill.
    staged_cell = sheet.cell(10, 6)
    assert staged_cell.value == 95
    assert staged_cell.fill.start_color.rgb == AMBER_FILL.start_color.rgb
    # Neighbouring unstaged cells keep their baseline values, no fill.
    assert sheet.cell(10, 5).value == "Flow Transmitter 101"
    assert sheet.cell(10, 5).fill.patternType is None
    assert sheet.cell(11, 6).value == 8
    assert sheet.cell(11, 6).fill.patternType is None
    book.close()


def test_export_clears_stale_rows(template):
    baselines = {"P_AIn": {"FIT_101": BASELINES["P_AIn"]["FIT_101"]}}
    book = _save_and_reload(build_export_workbook(template, baselines, [], "TestPLC"))
    sheet = book["P_AIn"]
    assert get_aoi_setup(sheet) == 1
    assert sheet.cell(10, 3).value == "FIT_101"
    # The template's second row is gone, not left behind as a stale instance.
    assert sheet.cell(11, 3).value is None
    assert sheet.cell(11, 5).value is None
    assert book["Setup"].cell(2, 4).value == 1
    book.close()


def test_export_leaves_unbaselined_aoi_untouched(template):
    book = _save_and_reload(build_export_workbook(template, {}, [], "TestPLC"))
    sheet = book["P_AIn"]
    # No baseline: the template's own rows survive untouched.
    assert get_aoi_setup(sheet) == 2
    assert sheet.cell(10, 3).value == "FIT_101"
    assert sheet.cell(11, 6).value == 8.0
    book.close()


def test_export_missing_template_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        build_export_workbook(tmp_path / "nope.xlsx", BASELINES, [], "TestPLC")
