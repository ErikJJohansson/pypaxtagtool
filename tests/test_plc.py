"""Tests for pypax.plc."""

import pytest

from fakes.fake_plc import FakePLC
from pypax.plc import (
    failed_tag_formatter,
    get_aoi_tag_instances,
    get_dim_list,
    get_failed_tags,
    make_tag_list,
    read_plc_row,
    write_plc_row,
)


@pytest.fixture
def plc():
    fake = FakePLC(commpath="10.0.0.1/0", plc_name="TestPLC")
    # Real LogixDriver.tags holds base instance tags; the fake seeds one
    # entry per subtag, so override with the real-driver shape here.
    fake.tags = {
        "FIT_101": {"data_type_name": "P_AIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
        "PIT_201": {"data_type_name": "P_AIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
        "XIC_301": {"data_type_name": "P_DIn", "dim": 0, "dimensions": [0, 0, 0], "alias": False},
        "ALIAS_1": {"data_type_name": "P_AIn", "dim": 0, "dimensions": [0, 0, 0], "alias": True},
    }
    fake.open()
    yield fake
    fake.close()


def test_get_aoi_tag_instances_filters_by_type_and_skips_aliases(plc):
    assert get_aoi_tag_instances(plc, "P_AIn") == ["FIT_101", "PIT_201"]
    assert get_aoi_tag_instances(plc, "P_DIn") == ["XIC_301"]
    assert get_aoi_tag_instances(plc, "P_Motor") == []


def test_get_aoi_tag_instances_expands_arrays(plc):
    plc.tags["ARR_1"] = {
        "data_type_name": "P_AIn",
        "dim": 1,
        "dimensions": [2, 0, 0],
        "alias": False,
    }
    assert get_aoi_tag_instances(plc, "P_AIn") == [
        "FIT_101",
        "PIT_201",
        "ARR_1[0]",
        "ARR_1[1]",
    ]


def test_get_dim_list_single_and_multi_dim():
    assert get_dim_list("TAG", [2]) == ["TAG[0]", "TAG[1]"]
    assert get_dim_list("TAG", [2, 3]) == [
        "TAG[0][0]",
        "TAG[0][1]",
        "TAG[0][2]",
        "TAG[1][0]",
        "TAG[1][1]",
        "TAG[1][2]",
    ]
    # Zero dimensions are dropped; a 3-D declaration with a zero still works.
    assert get_dim_list("TAG", [2, 0, 2]) == [
        "TAG[0][0]",
        "TAG[0][1]",
        "TAG[1][0]",
        "TAG[1][1]",
    ]


def test_get_dim_list_all_zero_dimensions_yields_base_tag():
    # Degenerate but real: filter leaves nothing, product yields one empty
    # tuple, so the bare base tag is returned.
    assert get_dim_list("TAG", [0, 0, 0]) == ["TAG"]


def test_make_tag_list():
    assert make_tag_list("FIT_101", [".Cfg_Desc", ".Cfg_HiLim"]) == [
        "FIT_101.Cfg_Desc",
        "FIT_101.Cfg_HiLim",
    ]
    assert make_tag_list("FIT_101", []) == []


def test_read_plc_row_coerces_bool_real_and_string(plc):
    tag_data, result = read_plc_row(
        plc, ["FIT_101.Cfg_Desc", "FIT_101.Cfg_HiLim", "FIT_101.HiHi.Cfg_Exists"]
    )
    assert tag_data == [
        ("FIT_101.Cfg_Desc", "Flow Transmitter 101"),
        ("FIT_101.Cfg_HiLim", 90),  # whole-number REAL -> int
        ("FIT_101.HiHi.Cfg_Exists", 1),  # BOOL -> int
    ]
    assert isinstance(tag_data[1][1], int)
    assert all(result)


def test_read_plc_row_real_rounding(plc):
    plc.write(("FIT_101.Cfg_HiLim", 42.123456789))
    tag_data, _ = read_plc_row(plc, ["FIT_101.Cfg_HiLim"])
    assert tag_data == [("FIT_101.Cfg_HiLim", 42.123457)]  # .6f rounding

    plc.write(("FIT_101.Cfg_HiLim", 0.0000001))
    tag_data, _ = read_plc_row(plc, ["FIT_101.Cfg_HiLim"])
    assert tag_data == [("FIT_101.Cfg_HiLim", 1e-07)]  # scientific form


def test_read_plc_row_reports_failures(plc):
    tag_data, result = read_plc_row(plc, ["FIT_101.Cfg_HiLim", "NOPE_999.Cfg_HiLim"])
    assert tag_data[0] == ("FIT_101.Cfg_HiLim", 90)
    assert tag_data[1] == ("NOPE_999.Cfg_HiLim", None)
    assert [bool(r) for r in result] == [True, False]


def test_write_plc_row_single_pair_returns_list(plc):
    results = write_plc_row(plc, [("FIT_101.Cfg_HiLim", 42.5)])
    assert len(results) == 1
    assert [bool(r) for r in results] == [True]
    assert plc.read("FIT_101.Cfg_HiLim").value == 42.5


def test_write_plc_row_multiple_and_failed(plc):
    results = write_plc_row(
        plc,
        [("FIT_101.Cfg_HiLim", 42.5), ("NOPE_999.Cfg_HiLim", 1.0)],
    )
    assert [bool(r) for r in results] == [True, False]


def test_read_plc_row_single_tag_is_normalised_to_a_row(plc):
    # pycomm3 returns a single Tag (not a list) for a one-tag read; the
    # legacy code iterated that Tag's fields. read_plc_row must still
    # return a one-entry row.
    tag_data, result = read_plc_row(plc, ["FIT_101.Cfg_HiLim"])
    assert tag_data == [("FIT_101.Cfg_HiLim", 90)]
    assert len(result) == 1
    assert all(result)


def test_get_failed_tags():
    tags = ["A", "B", "C", "D"]
    assert get_failed_tags(tags, [True, False, True, False]) == ["B", "D"]
    assert get_failed_tags(tags, [True, True, True, True]) == []
    assert get_failed_tags([], []) == []


def test_failed_tag_formatter():
    assert failed_tag_formatter(["A", "B"], True).startswith(
        "****WARNING****\nCANNOT WRITE TO\nA, B"
    )
    assert failed_tag_formatter(["A"], False).startswith("****WARNING****\nCANNOT READ FROM\nA")
