"""Helpers to build small workbooks in the real template layout.

Layout (mirrors ProcessLibraryOnlineConfigTool.xlsm):
- ``Setup`` sheet: header row 1, one row per AOI; col 4 = instance count,
  col 5 = number of config items, col 8 = worksheet tab name.
- AOI sheet (here ``P_AIn``): row 6 = data types, rows 7+8 = subtag
  (top + bottom), row 9 = defaults, instances from row 10, tag name in
  col C, values from col E.

Only the 3 subtags the FakePLC seeds are used, so the golden test stays
small while exercising the real read/write paths end to end.
"""

from __future__ import annotations

import openpyxl

# (top tag row 7, bottom tag row 8, data type row 6)
SUBTAGS = [
    (".Cfg_", "Desc", "STRING_40"),
    (".Cfg_", "HiLim", "REAL"),
    (".HiHi.Cfg_", "Exists", "BOOL"),
]

SUBTAG_NAMES = [top + bottom for top, bottom, _ in SUBTAGS]


def build_workbook(path, instance_rows=None):
    """Build a minimal template workbook at ``path``.

    ``instance_rows`` is a list of ``(tag_name, [value, ...])`` rows written
    from row 10; pass ``None``/``[]`` for a blank template (read flow).
    """
    book = openpyxl.Workbook()
    setup = book.active
    setup.title = "Setup"
    setup.append(
        [
            "Srl. No.",
            "Class",
            "CLX Array Base Tag Name\n(If Applicable)",
            "Max Instance",
            "Number Of Cfg Items",
            "OPC Server",
            "OPC Topic / Shortcut",
            "Worksheet Tab Name",
            "Origin Row",
        ]
    )
    setup.append(
        [1, "P_AIn", "P_AIn", 0, len(SUBTAGS), "FactoryTalk Gateway", "Shortcut", "P_AIn", 7]
    )

    sheet = book.create_sheet("P_AIn")
    sheet.cell(1, 2).value = "P_AIn: Analog Input"
    for col, (_top, _bottom, _type) in enumerate(SUBTAGS, start=5):
        sheet.cell(4, col).value = f"header {col}"
        sheet.cell(5, col).value = "Local"
    for col, (top, bottom, data_type) in enumerate(SUBTAGS, start=5):
        sheet.cell(6, col).value = data_type
        sheet.cell(7, col).value = top
        sheet.cell(8, col).value = bottom
    sheet.cell(8, 1).value = "Unit"
    sheet.cell(8, 2).value = "Tag Description"
    sheet.cell(8, 3).value = "TagName"
    sheet.cell(8, 4).value = "Instance"
    for col, (_top, _bottom, _type) in enumerate(SUBTAGS, start=5):
        sheet.cell(9, col).value = None

    for i, (tag_name, values) in enumerate(instance_rows or []):
        row = 10 + i
        sheet.cell(row, 3).value = tag_name
        for j, value in enumerate(values):
            sheet.cell(row, 5 + j).value = value

    book.save(path)
    book.close()
    return path
