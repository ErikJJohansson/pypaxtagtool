"""Spreadsheet layout constants and AOI sheet detection.

These constants describe the PlantPAX config spreadsheet layout
(``ProcessLibraryOnlineConfigTool.xlsm``) and are the single source of
truth for it: the CLI, the webapp core, and the tests all import them
from here instead of duplicating magic numbers.
"""

from __future__ import annotations

# AOI data sheets: instances start at this row, tag name in this column,
# values from this column onward.
START_ROW = 10
START_COL = 5
NAME_COL = 3

# Header rows of an AOI sheet: data type, subtag top half, subtag bottom half.
DATA_TYPE_ROW = 6
TOP_TAG_ROW = 7
BOTTOM_TAG_ROW = 8

# Setup sheet (first sheet of the workbook) columns.
SETUP_NUM_INSTANCES_COL = 4  # "Max Instance"
SETUP_NUM_SUBTAGS_COL = 5  # "Number Of Cfg Items"
SETUP_TAB_NAME_COL = 8  # "Worksheet Tab Name"

SETUP_SHEET_NAME = "Setup"


def get_aoi_list(excel_book):
    """Return the AOI sheet names in a workbook.

    PlantPAX AOI sheets have ``_`` as their second character (``P_AIn``,
    ``L_CPU``, ...); the ``SIF`` sheets are included too. Short or
    malformed sheet names are skipped instead of raising ``IndexError``.
    """
    aoi_list = []
    for sheet in excel_book.sheetnames:
        if len(sheet) >= 2 and sheet[1] == "_":
            aoi_list.append(sheet)
        elif sheet.startswith("SIF"):
            aoi_list.append(sheet)
    return aoi_list
