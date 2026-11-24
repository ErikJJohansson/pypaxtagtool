"""Tests for the FakePLC harness itself.

The fake is the foundation every later phase tests against, so it must be
proven here: reads, writes, read-back, unknown-tag failure, PLC naming,
and AOI discovery by data type. No network, no real PLC.
"""

import pytest

from fakes.fake_plc import FakePLC, FakeTag


@pytest.fixture
def plc():
    fake = FakePLC(commpath="10.0.0.1/0", plc_name="TestPLC")
    fake.open()
    yield fake
    fake.close()


def test_get_plc_name(plc):
    assert plc.get_plc_name() == "TestPLC"


def test_get_plc_name_requires_open_connection():
    fake = FakePLC()
    with pytest.raises(RuntimeError):
        fake.get_plc_name()


def test_read_known_tags_returns_values_types_and_truthy_results(plc):
    results = plc.read("FIT_101.Cfg_HiLim", "FIT_101.Cfg_Desc", "FIT_101.HiHi.Cfg_Exists")
    assert [r[0] for r in results] == [
        "FIT_101.Cfg_HiLim",
        "FIT_101.Cfg_Desc",
        "FIT_101.HiHi.Cfg_Exists",
    ]
    assert [r[1] for r in results] == [90.0, "Flow Transmitter 101", True]
    assert [r[2] for r in results] == ["REAL", "STRING", "BOOL"]
    assert all(results)
    assert all(r.error is None for r in results)


def test_read_result_is_indexable_like_pycomm3_tag(plc):
    (result,) = plc.read("PIT_201.Cfg_HiLim")
    assert isinstance(result, FakeTag)
    assert result[0] == "PIT_201.Cfg_HiLim"
    assert result[1] == 8.0
    assert result[2] == "REAL"
    assert tuple(result)[:3] == ("PIT_201.Cfg_HiLim", 8.0, "REAL")


def test_read_unknown_tag_fails_cleanly(plc):
    (result,) = plc.read("NOPE_999.Cfg_HiLim")
    assert not result
    assert result.value is None
    assert result.error == "Unknown tag"


def test_read_mixed_known_and_unknown_tags(plc):
    results = plc.read("FIT_101.Cfg_HiLim", "NOPE_999.Cfg_HiLim")
    assert [bool(r) for r in results] == [True, False]
    assert not all(results)


def test_write_then_read_back(plc):
    assert plc.write(("FIT_101.Cfg_HiLim", 42.5)) == [True]
    (result,) = plc.read("FIT_101.Cfg_HiLim")
    assert result.value == 42.5


def test_write_multiple_pairs_returns_one_bool_each(plc):
    assert plc.write(
        ("FIT_101.Cfg_HiLim", 42.5),
        ("FIT_101.Cfg_Desc", "Updated desc"),
        ("FIT_101.HiHi.Cfg_Exists", False),
    ) == [True, True, True]
    values = [
        r.value
        for r in plc.read("FIT_101.Cfg_HiLim", "FIT_101.Cfg_Desc", "FIT_101.HiHi.Cfg_Exists")
    ]
    assert values == [42.5, "Updated desc", False]


def test_write_unknown_tag_returns_false_and_changes_nothing(plc):
    assert plc.write(("NOPE_999.Cfg_HiLim", 1.0)) == [False]
    (result,) = plc.read("FIT_101.Cfg_HiLim")
    assert result.value == 90.0


def test_write_log_records_writes_in_order(plc):
    plc.write(("FIT_101.Cfg_HiLim", 42.5), ("PIT_201.Cfg_LoLim", 1.5))
    assert plc.write_log == [("FIT_101.Cfg_HiLim", 42.5), ("PIT_201.Cfg_LoLim", 1.5)]


def test_tags_dict_supports_aoi_discovery_by_data_type(plc):
    p_ain = sorted(
        name for name, definition in plc.tags.items() if definition["data_type_name"] == "P_AIn"
    )
    assert p_ain == sorted(
        [
            "FIT_101.Cfg_Desc",
            "FIT_101.Cfg_Label",
            "FIT_101.Cfg_Tag",
            "FIT_101.Cfg_EU",
            "FIT_101.Cfg_HiLim",
            "FIT_101.Cfg_LoLim",
            "FIT_101.Cfg_HiHiLim",
            "FIT_101.Cfg_LoLoLim",
            "FIT_101.HiHi.Cfg_Exists",
            "FIT_101.HiHi.Cfg_Severity",
            "PIT_201.Cfg_Desc",
            "PIT_201.Cfg_Label",
            "PIT_201.Cfg_Tag",
            "PIT_201.Cfg_EU",
            "PIT_201.Cfg_HiLim",
            "PIT_201.Cfg_LoLim",
            "PIT_201.Cfg_HiHiLim",
            "PIT_201.Cfg_LoLoLim",
            "PIT_201.HiHi.Cfg_Exists",
            "PIT_201.HiHi.Cfg_Severity",
        ]
    )
    assert all(
        not definition["alias"] and definition["dim"] == 0 for definition in plc.tags.values()
    )


def test_reset_restores_seed_values_and_clears_write_log(plc):
    plc.write(("FIT_101.Cfg_HiLim", 42.5))
    plc.reset()
    (result,) = plc.read("FIT_101.Cfg_HiLim")
    assert result.value == 90.0
    assert plc.write_log == []


def test_read_and_write_require_open_connection():
    fake = FakePLC()
    with pytest.raises(RuntimeError):
        fake.read("FIT_101.Cfg_HiLim")
    with pytest.raises(RuntimeError):
        fake.write(("FIT_101.Cfg_HiLim", 1.0))
