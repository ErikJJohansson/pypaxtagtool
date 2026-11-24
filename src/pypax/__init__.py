"""PyPaxTagTool — PlantPAX tag configuration tool.

The core library lives in this package. The legacy CLI (``pypax.py``
at the repo root) is a thin shim over :mod:`pypax.cli`, and the webapp
backend is built on the same core, so the two cannot drift apart.
"""

from __future__ import annotations

from .diff import REAL_TOLERANCE, row_differences, values_equal
from .layout import (
    BOTTOM_TAG_ROW,
    DATA_TYPE_ROW,
    NAME_COL,
    SETUP_NUM_INSTANCES_COL,
    SETUP_NUM_SUBTAGS_COL,
    SETUP_SHEET_NAME,
    SETUP_TAB_NAME_COL,
    START_COL,
    START_ROW,
    TOP_TAG_ROW,
    get_aoi_list,
)
from .plc import (
    MAX_TAGS_PER_READ,
    MAX_TAGS_PER_WRITE,
    build_tag_index,
    failed_tag_formatter,
    get_aoi_tag_instances,
    get_dim_list,
    get_failed_tags,
    make_tag_list,
    read_aoi_instances,
    read_plc_row,
    write_plc_row,
)
from .spreadsheet import (
    get_aoi_setup,
    get_subtag,
    get_subtag_list,
    read_data_sheet_row,
    search_value_in_col,
    set_num_instances,
    write_sheet_row,
)

__all__ = [
    "BOTTOM_TAG_ROW",
    "DATA_TYPE_ROW",
    "MAX_TAGS_PER_READ",
    "MAX_TAGS_PER_WRITE",
    "NAME_COL",
    "REAL_TOLERANCE",
    "SETUP_NUM_INSTANCES_COL",
    "SETUP_NUM_SUBTAGS_COL",
    "SETUP_SHEET_NAME",
    "SETUP_TAB_NAME_COL",
    "START_COL",
    "START_ROW",
    "TOP_TAG_ROW",
    "build_tag_index",
    "failed_tag_formatter",
    "get_aoi_list",
    "get_aoi_setup",
    "get_aoi_tag_instances",
    "get_dim_list",
    "get_failed_tags",
    "get_subtag",
    "get_subtag_list",
    "make_tag_list",
    "read_aoi_instances",
    "read_data_sheet_row",
    "read_plc_row",
    "row_differences",
    "search_value_in_col",
    "set_num_instances",
    "values_equal",
    "write_plc_row",
    "write_sheet_row",
]
