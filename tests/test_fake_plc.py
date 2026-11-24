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
    result = plc.read("PIT_201.Cfg_HiLim")
    assert isinstance(result, FakeTag)
    assert result[0] == "PIT_201.Cfg_HiLim"
    assert result[1] == 8.0
    assert result[2] == "REAL"
    assert result[3] is None
    assert tuple(result) == ("PIT_201.Cfg_HiLim", 8.0, "REAL", None)
    assert len(result) == 4
    # pycomm3's Tag is a 4-field NamedTuple (tag, value, type, error);
    # it unpacks as 4 fields, with .error as the 4th field and attribute.
    tag, value, type_, error = result
    assert (tag, value, type_, error) == ("PIT_201.Cfg_HiLim", 8.0, "REAL", None)
    assert result.error is None


def test_failed_read_result_is_4_fields_with_error(plc):
    result = plc.read("NOPE_999.Cfg_HiLim")
    assert len(result) == 4
    assert tuple(result) == ("NOPE_999.Cfg_HiLim", None, None, "Unknown tag")
    tag, value, type_, error = result
    assert (tag, value, type_, error) == ("NOPE_999.Cfg_HiLim", None, None, "Unknown tag")
    assert not result
    assert result.error == "Unknown tag"


def test_read_single_tag_returns_single_tag_not_list(plc):
    result = plc.read("FIT_101.Cfg_HiLim")
    assert isinstance(result, FakeTag)
    assert not isinstance(result, list)
    assert result.value == 90.0


def test_read_unknown_tag_fails_cleanly(plc):
    result = plc.read("NOPE_999.Cfg_HiLim")
    assert not result
    assert result.value is None
    assert result.error == "Unknown tag"


def test_read_mixed_known_and_unknown_tags(plc):
    results = plc.read("FIT_101.Cfg_HiLim", "NOPE_999.Cfg_HiLim")
    assert [bool(r) for r in results] == [True, False]
    assert not all(results)


def test_write_then_read_back(plc):
    result = plc.write(("FIT_101.Cfg_HiLim", 42.5))
    assert isinstance(result, FakeTag)
    assert not isinstance(result, list)
    assert result
    assert result.value == 42.5
    assert plc.read("FIT_101.Cfg_HiLim").value == 42.5


def test_write_multiple_pairs_returns_one_tag_each(plc):
    results = plc.write(
        ("FIT_101.Cfg_HiLim", 42.5),
        ("FIT_101.Cfg_Desc", "Updated desc"),
        ("FIT_101.HiHi.Cfg_Exists", False),
    )
    assert isinstance(results, list)
    assert [bool(r) for r in results] == [True, True, True]
    values = [
        r.value
        for r in plc.read("FIT_101.Cfg_HiLim", "FIT_101.Cfg_Desc", "FIT_101.HiHi.Cfg_Exists")
    ]
    assert values == [42.5, "Updated desc", False]


def test_write_single_pair_returns_single_tag_not_list(plc):
    result = plc.write(("FIT_101.Cfg_HiLim", 42.5))
    assert isinstance(result, FakeTag)
    assert not isinstance(result, list)


def test_write_unknown_tag_returns_falsy_tag_and_changes_nothing(plc):
    result = plc.write(("NOPE_999.Cfg_HiLim", 1.0))
    assert not result
    assert result.error == "Unknown tag"
    assert plc.read("FIT_101.Cfg_HiLim").value == 90.0


def test_write_log_records_writes_in_order(plc):
    plc.write(("FIT_101.Cfg_HiLim", 42.5), ("PIT_201.Cfg_LoLim", 1.5))
    assert plc.write_log == [("FIT_101.Cfg_HiLim", 42.5), ("PIT_201.Cfg_LoLim", 1.5)]


def test_tags_dict_supports_aoi_discovery_by_data_type(plc):
    # Phase 2: tags holds base-instance entries (the real-driver shape)
    # alongside the seeded member entries, so both appear here.
    p_ain = sorted(
        name for name, definition in plc.tags.items() if definition["data_type_name"] == "P_AIn"
    )
    assert p_ain == sorted(
        [
            "FIT_101",
            "PIT_201",
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
    assert plc.read("FIT_101.Cfg_HiLim").value == 90.0
    assert plc.write_log == []


def test_read_and_write_require_open_connection():
    fake = FakePLC()
    with pytest.raises(RuntimeError):
        fake.read("FIT_101.Cfg_HiLim")
    with pytest.raises(RuntimeError):
        fake.write(("FIT_101.Cfg_HiLim", 1.0))


def test_read_log_records_one_entry_per_read_call(plc):
    plc.read("FIT_101.Cfg_HiLim")
    plc.read("FIT_101.Cfg_HiLim", "PIT_201.Cfg_HiLim")
    assert plc.read_log == [
        ("FIT_101.Cfg_HiLim",),
        ("FIT_101.Cfg_HiLim", "PIT_201.Cfg_HiLim"),
    ]


def test_write_calls_records_one_entry_per_write_call(plc):
    plc.write(("FIT_101.Cfg_HiLim", 1.0))
    plc.write(("FIT_101.Cfg_HiLim", 2.0), ("PIT_201.Cfg_HiLim", 3.0))
    assert plc.write_calls == [
        (("FIT_101.Cfg_HiLim", 1.0),),
        (("FIT_101.Cfg_HiLim", 2.0), ("PIT_201.Cfg_HiLim", 3.0)),
    ]


def test_add_instance_seeds_base_and_member_tags(plc):
    plc.add_instance(
        "TT_999",
        "P_AIn",
        [(".Cfg_Desc", "STRING", "Test transmitter"), (".Cfg_HiLim", "REAL", 50.0)],
    )
    assert plc.tags["TT_999"]["data_type_name"] == "P_AIn"
    assert plc.read("TT_999.Cfg_Desc").value == "Test transmitter"
    assert plc.read("TT_999.Cfg_HiLim").value == 50.0


def test_reset_clears_call_logs_and_added_instances(plc):
    plc.add_instance("TT_999", "P_AIn", [(".Cfg_HiLim", "REAL", 50.0)])
    plc.read("FIT_101.Cfg_HiLim")
    plc.write(("FIT_101.Cfg_HiLim", 1.0))
    plc.reset()
    assert plc.read_log == []
    assert plc.write_calls == []
    assert plc.write_log == []
    assert "TT_999" not in plc.tags
    assert plc.read("FIT_101.Cfg_HiLim").value == 90.0
