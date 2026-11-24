"""Export endpoint tests: GET /api/export against the FakePLC."""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest
from fastapi.testclient import TestClient

from fakes.fake_plc import FakePLC
from helpers import SUBTAG_NAMES, build_workbook
from pypax.export import AMBER_FILL
from pypax.spreadsheet import get_aoi_setup, read_data_sheet_row
from pypax.web import create_app

PLC_NAME = "TestPLC"
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture()
def workbook(tmp_path):
    """Small template workbook: P_AIn sheet, FIT_101 + PIT_201 rows."""
    path = tmp_path / "template.xlsx"
    return build_workbook(
        path,
        instance_rows=[
            ("FIT_101", ["Flow Transmitter 101", 90.0, 1]),
            ("PIT_201", ["Pressure Transmitter 201", 8.0, 0]),
        ],
    )


@pytest.fixture()
def app(workbook, tmp_path):
    return create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(workbook),
        audit_dir=str(tmp_path / "audit"),
    )


@pytest.fixture()
def client(app):
    return TestClient(app)


@pytest.fixture()
def connected(client):
    response = client.post("/api/connect", json={"commpath": "10.0.0.1/0"})
    assert response.status_code == 200
    assert response.json()["connected"] is True
    return client


def _load_export(response):
    """Parse an export response back into an openpyxl workbook."""
    assert response.status_code == 200
    assert response.headers["content-type"].startswith(XLSX_MEDIA_TYPE)
    return openpyxl.load_workbook(BytesIO(response.content))


def test_export_returns_baseline_plus_staged(connected):
    # Load the baseline, then stage one change via paste.
    assert connected.get("/api/aois/P_AIn").status_code == 200
    paste = connected.post(
        "/api/aois/P_AIn/paste",
        json={
            "tsv": "95",
            "anchor_instance": "FIT_101",
            "anchor_subtag": ".Cfg_HiLim",
        },
    )
    assert paste.status_code == 200
    assert paste.json()["applied"] == 1

    response = connected.get("/api/export")
    assert response.headers["content-disposition"] == (
        'attachment; filename="TestPLC_ConfigTags.xlsx"'
    )
    book = _load_export(response)
    sheet = book["P_AIn"]
    assert get_aoi_setup(sheet) == 2
    base, tag_data = read_data_sheet_row(sheet, 10, SUBTAG_NAMES)
    assert base == "FIT_101"
    assert tag_data == [
        ("FIT_101.Cfg_Desc", "Flow Transmitter 101"),
        ("FIT_101.Cfg_HiLim", 95.0),
        ("FIT_101.HiHi.Cfg_Exists", 1),
    ]
    base, tag_data = read_data_sheet_row(sheet, 11, SUBTAG_NAMES)
    assert base == "PIT_201"
    assert tag_data[1] == ("PIT_201.Cfg_HiLim", 8.0)
    # The staged cell carries the amber fill; the rest do not.
    assert sheet.cell(10, 6).fill.start_color.rgb == AMBER_FILL.start_color.rgb
    assert sheet.cell(11, 6).fill.patternType is None
    assert book["Setup"].cell(2, 4).value == 2
    book.close()


def test_export_with_nothing_staged(connected):
    connected.get("/api/aois/P_AIn")
    book = _load_export(connected.get("/api/export"))
    sheet = book["P_AIn"]
    _base, tag_data = read_data_sheet_row(sheet, 10, SUBTAG_NAMES)
    assert tag_data[1] == ("FIT_101.Cfg_HiLim", 90.0)
    for row in (10, 11):
        for col in range(3, 8):
            assert sheet.cell(row, col).fill.patternType is None
    book.close()


def test_export_no_template_returns_503(tmp_path):
    app = create_app(
        plc_factory=lambda commpath: FakePLC(commpath),
        template_path=str(tmp_path / "nope.xlsx"),
        audit_dir=str(tmp_path / "audit"),
    )
    client = TestClient(app)
    assert client.post("/api/connect", json={"commpath": "10.0.0.1/0"}).status_code == 200
    response = client.get("/api/export")
    assert response.status_code == 503
    assert "template" in response.json()["detail"]


def test_export_disconnected_uses_fallback_filename(client, workbook):
    response = client.get("/api/export")
    assert response.status_code == 200
    assert response.headers["content-disposition"] == (
        'attachment; filename="PyPax_ConfigTags.xlsx"'
    )
    # No baseline while disconnected: the template's own rows come back.
    book = _load_export(response)
    sheet = book["P_AIn"]
    assert get_aoi_setup(sheet) == 2
    assert sheet.cell(10, 3).value == "FIT_101"
    assert sheet.cell(11, 6).value == 8.0
    book.close()
