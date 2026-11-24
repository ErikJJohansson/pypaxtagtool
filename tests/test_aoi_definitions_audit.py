"""Light PlantPAX v4.10.6 audit: definitions vs template spreadsheet.

Per Erik, 4.10.4 -> 4.10.6 is logic-only, no tag changes, so the
v4.10.4 AOI definitions should carry over unchanged. This test verifies
that instead of reworking anything: every AOI sheet's subtag list
(rows 7+8) must equal its ``AOI_definitions.py`` entry, and vice versa.

The one known exception is ``P_StrapTblCfg``: it has a definitions
entry but no sheet in the template. A second known exception is the
``P_ValveMO`` spelling below: the template's own description for that
column ("1=Coasting into LS when stoped changes state...") supports
the sheet spelling ``.Cfg_CoastToLS`` over the definitions file's
``.Cfg_CoastTolS``, but with no PLC or Studio 5000 export available
here that cannot be proven — and the CLI reads subtags from the
spreadsheet, never from ``AOI_definitions.py``, so the mismatch has
no runtime effect today. It needs a human check against a real
P_ValveMO AOI; if confirmed, the one-line fix is in
``AOI_definitions.py``. Any mismatch beyond the known ones fails.
"""

import importlib.util
from pathlib import Path

import openpyxl

from pypax.layout import get_aoi_list
from pypax.spreadsheet import get_subtag_list

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = REPO_ROOT / "ProcessLibraryOnlineConfigTool.xlsm"

# Spot-checks from the original manual verification (Nov 24, 2026).
SPOT_CHECK_COUNTS = {
    "P_AIn": 101,
    "P_Motor": 81,
    "P_PIDE": 185,
    "L_CPU": 6,
}

# Definitions entries known to have no sheet in the template.
KNOWN_SHEET_LESS_DEFINITIONS = {"P_StrapTblCfg"}

# Known subtag disagreements: {AOI: {"sheet_only": [...], "defs_only": [...]}}
KNOWN_MISMATCHES = {
    "P_ValveMO": {"sheet_only": [".Cfg_CoastToLS"], "defs_only": [".Cfg_CoastTolS"]},
}


def _load_definitions():
    spec = importlib.util.spec_from_file_location(
        "aoi_definitions", REPO_ROOT / "AOI_definitions.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.AOI_DEFINITIONS


def test_every_aoi_sheet_matches_its_definitions():
    definitions = _load_definitions()
    book = openpyxl.load_workbook(TEMPLATE, read_only=True, data_only=True)
    try:
        aoi_sheets = get_aoi_list(book)
        assert aoi_sheets, "no AOI sheets found in template"

        mismatches = {}
        for aoi in aoi_sheets:
            assert aoi in definitions, f"AOI sheet {aoi} has no definitions entry"
            sheet_subtags = get_subtag_list(book[aoi])
            if sheet_subtags != definitions[aoi]:
                mismatches[aoi] = {
                    "sheet_only": sorted(set(sheet_subtags) - set(definitions[aoi])),
                    "defs_only": sorted(set(definitions[aoi]) - set(sheet_subtags)),
                }
        assert mismatches == KNOWN_MISMATCHES, f"unexpected subtag mismatches: {mismatches}"

        for aoi, count in SPOT_CHECK_COUNTS.items():
            assert len(definitions[aoi]) == count, f"{aoi} should have {count} subtags"
    finally:
        book.close()


def test_definitions_without_sheets_are_only_the_known_case():
    definitions = _load_definitions()
    book = openpyxl.load_workbook(TEMPLATE, read_only=True, data_only=True)
    try:
        sheetless = set(definitions) - set(get_aoi_list(book))
    finally:
        book.close()
    assert sheetless == KNOWN_SHEET_LESS_DEFINITIONS, (
        f"unexpected definitions without sheets: {sheetless - KNOWN_SHEET_LESS_DEFINITIONS}"
    )
