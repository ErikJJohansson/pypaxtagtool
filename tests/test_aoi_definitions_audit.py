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
import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

from pypax.aoi import dump_aoi_definitions, load_aoi_definitions
from pypax.layout import get_aoi_list
from pypax.spreadsheet import get_subtag_list

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = REPO_ROOT / "ProcessLibraryOnlineConfigTool.xlsm"
AOI_YAML = REPO_ROOT / "config" / "aoi" / "v4.10.6.yaml"
EXTRACTOR = REPO_ROOT / "scripts" / "extract_aoi_definitions.py"

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


def test_yaml_matches_spreadsheet_and_definitions():
    """The generated v4.10.6 YAML is the spreadsheet's subtag lists.

    It matches the spreadsheet exactly, and matches AOI_definitions.py
    except for the known P_ValveMO spelling disagreement; the YAML has
    no P_StrapTblCfg entry because there is no such sheet.
    """
    yaml_defs = load_aoi_definitions(AOI_YAML)
    definitions = _load_definitions()
    book = openpyxl.load_workbook(TEMPLATE, read_only=True, data_only=True)
    try:
        aoi_sheets = get_aoi_list(book)
        assert set(yaml_defs) == set(aoi_sheets)
        for aoi in aoi_sheets:
            assert yaml_defs[aoi] == get_subtag_list(book[aoi]), f"{aoi} YAML != spreadsheet"
    finally:
        book.close()
    assert "P_StrapTblCfg" not in yaml_defs
    mismatches = {}
    for aoi, yaml_subtags in yaml_defs.items():
        if yaml_subtags != definitions[aoi]:
            mismatches[aoi] = {
                "sheet_only": sorted(set(yaml_subtags) - set(definitions[aoi])),
                "defs_only": sorted(set(definitions[aoi]) - set(yaml_subtags)),
            }
    assert mismatches == KNOWN_MISMATCHES, f"unexpected YAML vs definitions diffs: {mismatches}"


def test_extractor_regenerates_the_checked_in_yaml(tmp_path):
    """Running the extractor reproduces the checked-in file byte for byte."""
    regenerated = tmp_path / "v4.10.6.yaml"
    result = subprocess.run(
        [sys.executable, str(EXTRACTOR), "--output", str(regenerated)],
        capture_output=True,
        text=True,
        cwd=REPO_ROOT,
    )
    assert result.returncode == 0, result.stderr
    assert regenerated.read_text(encoding="utf-8") == AOI_YAML.read_text(encoding="utf-8")


def test_aoi_yaml_round_trip(tmp_path):
    loaded = load_aoi_definitions(AOI_YAML)
    assert len(loaded) == 56
    assert len(loaded["P_AIn"]) == 101
    rewritten = tmp_path / "roundtrip.yaml"
    dump_aoi_definitions(loaded, rewritten)
    assert load_aoi_definitions(rewritten) == loaded
    assert rewritten.read_text(encoding="utf-8") == AOI_YAML.read_text(encoding="utf-8")


def test_aoi_yaml_loader_rejects_bad_shapes(tmp_path):
    bad = tmp_path / "bad.yaml"
    bad.write_text("P_AIn:\n  - '.Cfg_Desc'\nP_AIn:\n  - '.Cfg_HiLim'\n", encoding="utf-8")
    with pytest.raises(ValueError, match="duplicate"):
        load_aoi_definitions(bad)
    bad.write_text("not a mapping\n", encoding="utf-8")
    with pytest.raises(ValueError):
        load_aoi_definitions(bad)
    bad.write_text("P_AIn:\n  - .Cfg_Desc\n", encoding="utf-8")
    # Plain (unquoted) scalars are valid YAML and load fine.
    assert load_aoi_definitions(bad) == {"P_AIn": [".Cfg_Desc"]}
    bad.write_text("P_AIn:\n  - 42\n", encoding="utf-8")
    with pytest.raises(ValueError, match="list of subtag strings"):
        load_aoi_definitions(bad)
