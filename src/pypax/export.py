"""Spreadsheet export: write the webapp's PLC snapshot back to ``.xlsx``.

The export round-trips with the CLI: it starts from the template
workbook (``ProcessLibraryOnlineConfigTool.xlsm``), rewrites each AOI
sheet's instance rows from the webapp's PLC baseline — tag name in
col C, values from col E, same layout the CLI reads and writes — and
applies staged (unwritten) changes on top, marked with an amber fill
so a user opening the file can see what is staged but not yet in the
PLC. Setup-sheet instance counts are updated to match.

The exported file is a plain ``.xlsx`` (no macros — the CLI reads it
with ``keep_vba=False`` anyway) and its name carries the PLC name
(``<plc_name>_ConfigTags.xlsx``), which the CLI's write mode requires:
it refuses a file whose name does not contain the PLC name.

Only AOIs with a baseline are rewritten; an AOI with no baseline
(e.g. an export while disconnected) keeps the template's rows as-is.
"""

from __future__ import annotations

import openpyxl
from openpyxl.styles import PatternFill

from .layout import (
    NAME_COL,
    SETUP_SHEET_NAME,
    START_COL,
    START_ROW,
    get_aoi_list,
)
from .spreadsheet import get_aoi_setup, get_subtag_list, set_num_instances, write_sheet_row

# Amber fill for staged-but-unwritten cells (same "changed" meaning as
# the frontend's amber highlighting). Full 8-digit ARGB: a 6-digit hex
# would be stored with alpha 00 (transparent) by openpyxl.
AMBER_FILL = PatternFill(start_color="FFFFC000", end_color="FFFFC000", fill_type="solid")
_NO_FILL = PatternFill()


def _clear_instance_rows(sheet, subtag_count):
    """Clear values (and any fill) of the sheet's current instance rows.

    Only the consecutive named rows from ``START_ROW`` — the rows the
    CLI counts via :func:`get_aoi_setup` — in the name + value columns.
    Header rows (4–9) and defaults are untouched.
    """
    for row in range(START_ROW, START_ROW + get_aoi_setup(sheet)):
        for col in range(NAME_COL, START_COL + subtag_count):
            cell = sheet.cell(row, col)
            cell.value = None
            cell.fill = _NO_FILL


def build_export_workbook(template_path, baselines, staged_changes, plc_name=None):
    """Build the export workbook from the template.

    ``baselines`` maps AOI name -> ``{instance: {subtag: value}}`` (the
    last-read PLC values); ``staged_changes`` is an iterable of change
    dicts with ``aoi`` / ``instance`` / ``subtag`` / ``new_value`` keys.
    ``plc_name`` is accepted for the caller's filename (the CLI write
    mode requires it in the file name); no PLC-name cell exists in the
    template layout, so it is not written into the workbook itself.

    Returns the open workbook; the caller saves and closes it. Raises
    ``FileNotFoundError`` when the template cannot be found.
    """
    book = openpyxl.load_workbook(template_path, keep_vba=False)
    try:
        setup = book[SETUP_SHEET_NAME] if SETUP_SHEET_NAME in book.sheetnames else None
        staged_by_cell = {
            (change["aoi"], change["instance"], change["subtag"]): change
            for change in staged_changes
        }
        for aoi in get_aoi_list(book):
            sheet = book[aoi]
            sub_tags = get_subtag_list(sheet)
            rows = baselines.get(aoi)
            if rows is None:
                # No baseline for this AOI: leave the sheet exactly as
                # the template has it (e.g. an export while disconnected).
                continue
            _clear_instance_rows(sheet, len(sub_tags))
            for i, instance in enumerate(rows):
                values = rows[instance]
                tag_data = [(instance + sub, values.get(sub)) for sub in sub_tags]
                write_sheet_row(sheet, START_ROW + i, instance, tag_data)
                # Staged changes overwrite the baseline value in the cell
                # and are marked with the amber fill.
                for j, sub_tag in enumerate(sub_tags):
                    change = staged_by_cell.get((aoi, instance, sub_tag))
                    if change is not None:
                        cell = sheet.cell(START_ROW + i, START_COL + j)
                        cell.value = change["new_value"]
                        cell.fill = AMBER_FILL
            if setup is not None:
                set_num_instances(setup, aoi, len(rows))
        return book
    except Exception:
        book.close()
        raise
