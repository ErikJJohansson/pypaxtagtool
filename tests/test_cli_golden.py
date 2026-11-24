"""Golden test: the CLI's observable behaviour, before and after refactor.

Captures the CURRENT behaviour of ``pypax.py`` on a small generated
workbook (real template layout, 2 P_AIn instances x 3 subtags) driven
against the FakePLC. This test is written against the old, unrefactored
code first and must keep passing unchanged through the whole Phase 1
refactor — if it goes red, behaviour changed.

Note on the fake: a real pycomm3 ``LogixDriver.tags`` holds base instance
tags (``FIT_101`` with ``data_type_name == "P_AIn"``), while the Phase 0
FakePLC seeds one entry per *subtag*. The fixture below overrides
``fake.tags`` with base-instance entries, which is the real-driver shape
``get_aoi_tag_instances`` relies on; reads/writes still go through the
seeded subtag values.
"""

import sys

import openpyxl
import pytest

import pypax.cli as pypax_cli
from fakes.fake_plc import FakePLC
from helpers import SUBTAG_NAMES, build_workbook
from pypax.cli import main as cli_main

BASE_INSTANCE_TAGS = {
    "FIT_101": {"data_type_name": "P_AIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
    "PIT_201": {"data_type_name": "P_AIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
    "XIC_301": {"data_type_name": "P_DIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
}


@pytest.fixture
def fake(monkeypatch, tmp_path):
    plc = FakePLC(commpath="10.0.0.1/0", plc_name="TestPLC")
    plc.tags = dict(BASE_INSTANCE_TAGS)
    plc.open()
    monkeypatch.chdir(tmp_path)
    # cli.main() builds LogixDriver(commpath, init_tags=..., ...); hand it the fake.
    monkeypatch.setattr(pypax_cli, "LogixDriver", lambda *args, **kwargs: plc)
    yield plc
    plc.close()


def run_cli(monkeypatch, *argv):
    monkeypatch.setattr(sys, "argv", ["pypax.py", *argv])
    cli_main()


def test_read_flow_writes_plc_values_into_new_workbook(monkeypatch, tmp_path, fake, capsys):
    build_workbook(tmp_path / "template.xlsx")
    run_cli(monkeypatch, "10.0.0.1/0", "read", "template.xlsx")
    out = capsys.readouterr().out
    assert "Connected to TestPLC PLC at 10.0.0.1/0" in out
    assert "Finished reading from TestPLC PLC." in out

    outfile = tmp_path / "TestPLC_ConfigTags.xlsx"
    assert outfile.exists()
    book = openpyxl.load_workbook(outfile, data_only=True)
    assert book["Setup"].cell(2, 4).value == 2  # instance count on Setup sheet

    sheet = book["P_AIn"]
    assert [sheet.cell(10, 3).value, sheet.cell(11, 3).value] == ["FIT_101", "PIT_201"]
    # REAL 90.0 reads back as int 90 (REAL .is_integer() rule); BOOL as 1/0.
    assert [sheet.cell(10, 5).value, sheet.cell(10, 6).value, sheet.cell(10, 7).value] == [
        "Flow Transmitter 101",
        90,
        1,
    ]
    assert [sheet.cell(11, 5).value, sheet.cell(11, 6).value, sheet.cell(11, 7).value] == [
        "Pressure Transmitter 201",
        8,
        0,
    ]
    book.close()


def test_write_flow_writes_only_differences_then_reports_no_differences(
    monkeypatch, tmp_path, fake, capsys
):
    build_workbook(tmp_path / "template.xlsx")
    run_cli(monkeypatch, "10.0.0.1/0", "read", "template.xlsx")
    capsys.readouterr()

    # Edit two cells in the read-back file, leave everything else identical.
    outfile = tmp_path / "TestPLC_ConfigTags.xlsx"
    book = openpyxl.load_workbook(outfile)
    book["P_AIn"].cell(10, 6).value = 42.5  # FIT_101 HiLim 90 -> 42.5
    book["P_AIn"].cell(11, 5).value = "Updated desc"  # PIT_201 Desc
    book.save(outfile)
    book.close()

    run_cli(monkeypatch, "10.0.0.1/0", "write", "TestPLC_ConfigTags.xlsx")
    out = capsys.readouterr().out
    assert "Writing 2 tag changes to instances of P_AIn" in out
    assert "Finished writing to TestPLC PLC." in out
    assert fake.write_log == [
        ("FIT_101.Cfg_HiLim", 42.5),
        ("PIT_201.Cfg_Desc", "Updated desc"),
    ]
    # main() closes the PLC when it finishes; reopen for the read-back check.
    fake.open()
    assert [r.value for r in fake.read("FIT_101.Cfg_HiLim", "PIT_201.Cfg_Desc")] == [
        42.5,
        "Updated desc",
    ]

    # Second write: nothing changed, so nothing is written.
    run_cli(monkeypatch, "10.0.0.1/0", "write", "TestPLC_ConfigTags.xlsx")
    out = capsys.readouterr().out
    assert "No differences for instances of P_AIn" in out
    assert fake.write_log == [
        ("FIT_101.Cfg_HiLim", 42.5),
        ("PIT_201.Cfg_Desc", "Updated desc"),
    ]


def test_write_mode_rejects_filename_without_plc_name(monkeypatch, tmp_path, fake):
    build_workbook(tmp_path / "template.xlsx")
    with pytest.raises(SystemExit):
        run_cli(monkeypatch, "10.0.0.1/0", "write", "template.xlsx")


def test_subtag_layout_matches_fake_seed():
    # Guard: if the helper's subtags drift from the fake's seed tags, the
    # golden test would silently exercise the wrong tags.
    assert SUBTAG_NAMES == [".Cfg_Desc", ".Cfg_HiLim", ".HiHi.Cfg_Exists"]
