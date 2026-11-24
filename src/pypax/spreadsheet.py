"""Reading/writing AOI data in the spreadsheet.

Moved verbatim in behaviour from the legacy ``pypax.py`` CLI, except
that scans which used to touch one cell at a time now iterate rows
(``iter_rows``), which also works on read-only workbooks — the mode
write mode opens files in.
"""

from __future__ import annotations

from .layout import (
    BOTTOM_TAG_ROW,
    DATA_TYPE_ROW,
    NAME_COL,
    SETUP_NUM_INSTANCES_COL,
    SETUP_TAB_NAME_COL,
    START_COL,
    START_ROW,
    TOP_TAG_ROW,
)


def get_subtag(sheet, column):
    """Subtag for one column: top half (row 7) + bottom half (row 8)."""
    return str(sheet.cell(TOP_TAG_ROW, column).value) + str(
        sheet.cell(BOTTOM_TAG_ROW, column).value
    )


def get_subtag_list(sheet):
    """All subtags of an AOI sheet, in column order.

    Columns run from ``START_COL`` until the first column whose top and
    bottom halves are both blank (``'NoneNone'``).
    """
    top_row, bottom_row = (
        tuple(row)
        for row in sheet.iter_rows(min_row=TOP_TAG_ROW, max_row=BOTTOM_TAG_ROW, values_only=True)
    )
    sub_tag_list = []
    for top, bottom in zip(top_row[START_COL - 1 :], bottom_row[START_COL - 1 :]):
        sub_tag = str(top) + str(bottom)
        if sub_tag == "NoneNone":
            break
        sub_tag_list.append(sub_tag)
    return sub_tag_list


def search_value_in_col(sheet, search_string, col_idx=1):
    """Return the first row whose ``col_idx`` cell equals ``search_string``."""
    for row_idx, (value,) in enumerate(
        sheet.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True), start=1
    ):
        if value == search_string:
            return row_idx
    return None


def get_aoi_setup(sheet):
    """Number of AOI instances listed in a sheet (names from ``START_ROW``).

    Counts consecutive non-blank names in the name column; the first blank
    name ends the list. A blank string counts as a name, as before.
    """
    count = 0
    for (cell_value,) in sheet.iter_rows(
        min_row=START_ROW, min_col=NAME_COL, max_col=NAME_COL, values_only=True
    ):
        if cell_value is None:
            break
        count += 1
    return count


def set_num_instances(sheet, aoi_name, num):
    """Write the instance count for ``aoi_name`` on the Setup sheet."""
    aoi_row = search_value_in_col(sheet, aoi_name, SETUP_TAB_NAME_COL)
    if aoi_row is not None:
        sheet.cell(aoi_row, SETUP_NUM_INSTANCES_COL).value = num


def write_sheet_row(sheet, row, base_tag, tag_data):
    """Write one instance row: tag name plus one value per subtag."""
    sheet.cell(row, NAME_COL).value = base_tag
    for i in range(len(tag_data)):
        sheet.cell(row, START_COL + i).value = tag_data[i][1]


def read_data_sheet_row(sheet, row, sub_tags):
    """Read one instance row as ``(base_tag, [(full_tag, value), ...])``.

    Values are coerced by the data-type row, exactly as the legacy CLI
    did: INT/BOOL to ``int``, REAL to ``float``, STRING kept as-is, and a
    blank cell becomes ``''``.
    """
    base_tag = sheet.cell(row, NAME_COL).value
    if not sub_tags:
        return base_tag, []

    (values,) = sheet.iter_rows(
        min_row=row,
        max_row=row,
        min_col=START_COL,
        max_col=START_COL + len(sub_tags) - 1,
        values_only=True,
    )
    (types,) = sheet.iter_rows(
        min_row=DATA_TYPE_ROW,
        max_row=DATA_TYPE_ROW,
        min_col=START_COL,
        max_col=START_COL + len(sub_tags) - 1,
        values_only=True,
    )

    tag_data = []
    for i, sub_tag in enumerate(sub_tags):
        cell_value = values[i]
        cell_type = types[i]
        if cell_value is None:
            subtag_value = (base_tag + sub_tag, "")
        elif "INT" in cell_type:
            subtag_value = (base_tag + sub_tag, int(cell_value))
        elif "BOOL" in cell_type:
            subtag_value = (base_tag + sub_tag, int(cell_value))
        elif "STRING" in cell_type:
            subtag_value = (base_tag + sub_tag, cell_value)
        elif "REAL" in cell_type:
            subtag_value = (base_tag + sub_tag, float(cell_value))
        else:
            subtag_value = (base_tag + sub_tag, cell_value)
        tag_data.append(subtag_value)

    return base_tag, tag_data
