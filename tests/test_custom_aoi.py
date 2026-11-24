"""Tests for the custom-AOI store (``config/aoi/custom/*.yaml``)."""

from datetime import date

import pytest

from pypax.aoi import load_aoi_definitions
from pypax.custom_aoi import (
    check_collision,
    load_custom_aois,
    save_custom_aoi,
)
from pypax.l5x import Column

COLUMNS = [
    Column(subtag=".Cfg_Desc", data_type="STRING_40", description="Desc", default="Valve"),
    Column(subtag=".Cfg_HasStop", data_type="BOOL", description="", default=False),
    Column(subtag=".MyAlarm.Cfg_Severity", data_type="DINT", description="", default=750),
]


def test_save_and_load_round_trip(tmp_path):
    target = save_custom_aoi(
        "P_TestValve", COLUMNS, "P_TestValve.L5X", "1.02.00 Test", path=tmp_path
    )
    assert target == tmp_path / "P_TestValve.yaml"
    text = target.read_text(encoding="utf-8")
    assert "P_TestValve.L5X" in text
    assert "1.02.00 Test" in text
    assert date.today().isoformat() in text
    # Same shape as the generated library YAML: the existing loader
    # reads it back unchanged.
    assert load_aoi_definitions(target) == {
        "P_TestValve": [".Cfg_Desc", ".Cfg_HasStop", ".MyAlarm.Cfg_Severity"]
    }
    assert load_custom_aois(tmp_path) == {
        "P_TestValve": [".Cfg_Desc", ".Cfg_HasStop", ".MyAlarm.Cfg_Severity"]
    }


def test_save_sanitizes_newlines_in_header_fields(tmp_path):
    """A newline in the source name or revision must not inject YAML.

    The header interpolates both into ``#`` comment lines; a crafted
    file name (e.g. a scan-dir file with a newline in its name) could
    otherwise smuggle a second AOI definition into the YAML file,
    bypassing name validation and collision checks.
    """
    evil_name = "A\nP_Evil:\n  - '.Cfg_X'\n#Z.L5X"
    evil_revision = "1.0\nP_Evil2:\n  - '.Cfg_Y'"
    save_custom_aoi("P_Good", COLUMNS, evil_name, evil_revision, path=tmp_path)
    definitions = load_custom_aois(tmp_path)
    assert definitions == {"P_Good": [c.subtag for c in COLUMNS]}
    text = (tmp_path / "P_Good.yaml").read_text(encoding="utf-8")
    # The collapsed header may still mention the names inside a comment,
    # but no YAML line may define them as AOIs.
    assert "\nP_Evil:" not in text
    assert "\nP_Evil2:" not in text


def test_save_accepts_plain_subtag_strings(tmp_path):
    target = save_custom_aoi(
        "P_Plain", [".Cfg_Desc", ".Cfg_Tag"], "P_Plain.L5X", "1.0", path=tmp_path
    )
    assert load_aoi_definitions(target) == {"P_Plain": [".Cfg_Desc", ".Cfg_Tag"]}


def test_save_to_exact_file_path(tmp_path):
    target = tmp_path / "elsewhere" / "P_File.yaml"
    assert save_custom_aoi("P_File", COLUMNS, "P_File.L5X", "1.0", path=target) == target
    assert target.exists()


def test_save_rejects_bad_names_and_columns(tmp_path):
    for bad_name in ("", "P Valve", "../P_Evil", "1ABC", "P-Valve"):
        with pytest.raises(ValueError, match="AOI name"):
            save_custom_aoi(bad_name, COLUMNS, "x.L5X", "1.0", path=tmp_path)
    with pytest.raises(ValueError, match="duplicate subtags"):
        save_custom_aoi(
            "P_Dup",
            [".Cfg_Desc", ".Cfg_Desc"],
            "x.L5X",
            "1.0",
            path=tmp_path,
        )
    with pytest.raises(ValueError, match="not safe for YAML"):
        save_custom_aoi(
            "P_Bad",
            [".Cfg_Desc\ninjected: true"],
            "x.L5X",
            "1.0",
            path=tmp_path,
        )


def test_load_custom_aois_merges_and_rejects_duplicates(tmp_path):
    save_custom_aoi("P_One", [".Cfg_Desc"], "P_One.L5X", "1.0", path=tmp_path)
    save_custom_aoi("P_Two", [".Cfg_Tag"], "P_Two.L5X", "1.0", path=tmp_path)
    assert load_custom_aois(tmp_path) == {
        "P_One": [".Cfg_Desc"],
        "P_Two": [".Cfg_Tag"],
    }
    # A second file defining the same AOI is rejected.
    (tmp_path / "P_One_again.yaml").write_text("P_One:\n  - '.Cfg_Tag'\n", encoding="utf-8")
    with pytest.raises(ValueError, match="defined in both"):
        load_custom_aois(tmp_path)


def test_load_custom_aois_missing_dir_returns_empty(tmp_path):
    assert load_custom_aois(tmp_path / "nope") == {}


def test_check_collision_against_template_and_custom(tmp_path):
    # The real template sheet list: P_ValveMO is a library AOI.
    with pytest.raises(ValueError, match="template spreadsheet"):
        check_collision("P_ValveMO", custom_dir=tmp_path)
    # A fresh name passes against the real template and an empty dir.
    check_collision("P_CustomThing", custom_dir=tmp_path)
    save_custom_aoi("P_CustomThing", COLUMNS, "P_CustomThing.L5X", "1.0", path=tmp_path)
    with pytest.raises(ValueError, match="already exists as a custom AOI"):
        check_collision("P_CustomThing", template_path=None, custom_dir=tmp_path)
    # Case-insensitive: ControlLogix types are case-insensitive, and
    # on a case-insensitive filesystem the YAML files would collide.
    with pytest.raises(ValueError, match="template spreadsheet"):
        check_collision("p_valvemo", custom_dir=tmp_path)
    with pytest.raises(ValueError, match="already exists as a custom AOI"):
        check_collision("p_customthing", template_path=None, custom_dir=tmp_path)
    # Skipping both sources never collides.
    check_collision("P_ValveMO", template_path=None, custom_dir=None)


def test_save_to_default_location_refuses_library_name():
    # Raises before writing anything: no side effect on the repo.
    with pytest.raises(ValueError, match="template spreadsheet"):
        save_custom_aoi("P_ValveMO", COLUMNS, "P_ValveMO.L5X", "4.10.06")
