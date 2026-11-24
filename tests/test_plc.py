"""Tests for pypax.plc."""

import pytest

from fakes.fake_plc import FakePLC
from pypax.plc import (
    failed_tag_formatter,
    get_aoi_tag_instances,
    get_dim_list,
    get_failed_tags,
    make_tag_list,
    read_aoi_instances,
    read_plc_row,
    write_plc_row,
)


@pytest.fixture
def plc():
    fake = FakePLC(commpath="10.0.0.1/0", plc_name="TestPLC")
    fake.open()
    yield fake
    fake.close()


def test_get_aoi_tag_instances_filters_by_type_and_skips_aliases(plc):
    plc.tags["ALIAS_1"] = {
        "data_type_name": "P_AIn",
        "dim": 0,
        "dimensions": [0, 0, 0],
        "alias": True,
    }
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


def test_get_aoi_tag_instances_ignores_member_paths(plc):
    # The fake's tags dict also holds member entries (FIT_101.Cfg_Desc
    # with the AOI data type); discovery must return base instances only.
    assert "FIT_101.Cfg_Desc" not in get_aoi_tag_instances(plc, "P_AIn")
    assert get_aoi_tag_instances(plc, "P_AIn") == ["FIT_101", "PIT_201"]


def test_get_aoi_tag_instances_keeps_program_scoped_bases(plc):
    plc.tags["Program:MainProg.PT_500"] = {
        "data_type_name": "P_AIn",
        "dim": 0,
        "dimensions": [0, 0, 0],
        "alias": False,
    }
    plc.tags["Program:MainProg.PT_500.Cfg_HiLim"] = {
        "data_type_name": "P_AIn",
        "dim": 0,
        "dimensions": [0, 0, 0],
        "alias": False,
    }
    instances = get_aoi_tag_instances(plc, "P_AIn")
    assert "Program:MainProg.PT_500" in instances
    assert "Program:MainProg.PT_500.Cfg_HiLim" not in instances


def test_build_tag_index_groups_by_type_and_skips_aliases_and_members(plc):
    from pypax.plc import build_tag_index

    index = build_tag_index(plc)
    assert sorted(index) == ["P_AIn", "P_DIn"]
    assert [tag for tag, _ in index["P_AIn"]] == ["FIT_101", "PIT_201"]
    assert [tag for tag, _ in index["P_DIn"]] == ["XIC_301"]
    # Member paths never land in the index, even with the AOI data type.
    assert all("." not in tag for tags in index.values() for tag, _ in tags)


def test_get_aoi_tag_instances_accepts_a_prebuilt_index(plc):
    from pypax.plc import build_tag_index

    index = build_tag_index(plc)
    assert get_aoi_tag_instances(plc, "P_AIn", index=index) == ["FIT_101", "PIT_201"]
    assert get_aoi_tag_instances(plc, "P_DIn", index=index) == ["XIC_301"]


def test_tag_index_is_cached_and_rebuilt_when_tags_change(plc):
    get_aoi_tag_instances(plc, "P_AIn")
    cached = plc._pypax_tag_index
    assert cached is not None
    get_aoi_tag_instances(plc, "P_DIn")
    assert plc._pypax_tag_index is cached  # reused, not rebuilt

    plc.tags["TT_900"] = {
        "data_type_name": "P_AIn",
        "dim": 0,
        "dimensions": [0, 0, 0],
        "alias": False,
    }
    assert get_aoi_tag_instances(plc, "P_AIn") == ["FIT_101", "PIT_201", "TT_900"]
    assert plc._pypax_tag_index is not cached  # size changed -> rebuilt


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


def test_read_plc_row_failed_real_read_reports_none_instead_of_crashing(plc):
    # A REAL tag whose read failed still reports type REAL with value
    # None; the old code crashed on float(format(None, ".6e")).
    plc._values["FIT_101.Cfg_HiLim"] = None
    tag_data, result = read_plc_row(plc, ["FIT_101.Cfg_HiLim"])
    assert tag_data == [("FIT_101.Cfg_HiLim", None)]
    assert [bool(r) for r in result] == [False]


def test_read_plc_row_failed_bool_read_reports_none_instead_of_crashing(plc):
    # Same for BOOL: int(None) used to raise TypeError.
    plc._values["FIT_101.HiHi.Cfg_Exists"] = None
    tag_data, result = read_plc_row(plc, ["FIT_101.HiHi.Cfg_Exists"])
    assert tag_data == [("FIT_101.HiHi.Cfg_Exists", None)]
    assert [bool(r) for r in result] == [False]


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


def test_write_plc_row_chunks_large_writes(plc):
    for i in range(5):
        plc.add_instance(
            f"TT_{i}",
            "P_AIn",
            [(f".Cfg_Sub{j:02d}", "REAL", 0.0) for j in range(50)],
        )
    tag_data = [(f"TT_{i}.Cfg_Sub{j:02d}", float(i * 100 + j)) for i in range(5) for j in range(50)]
    results = write_plc_row(plc, tag_data)
    assert len(results) == 250
    assert all(results)
    # 250 pairs in chunks of at most 100: three plc.write calls, 100/100/50.
    assert [len(call) for call in plc.write_calls] == [100, 100, 50]
    assert plc.read("TT_0.Cfg_Sub00").value == 0.0
    assert plc.read("TT_4.Cfg_Sub49").value == 449.0


def test_write_plc_row_preserves_order_across_chunks(plc):
    results = write_plc_row(
        plc,
        [("FIT_101.Cfg_HiLim", 1.0), ("PIT_201.Cfg_HiLim", 2.0), ("NOPE_999.Cfg_X", 3.0)],
        max_tags_per_write=2,
    )
    assert [bool(r) for r in results] == [True, True, False]
    assert [len(call) for call in plc.write_calls] == [2, 1]
    assert plc.write_log == [
        ("FIT_101.Cfg_HiLim", 1.0),
        ("PIT_201.Cfg_HiLim", 2.0),
    ]


def test_write_plc_row_empty_writes_nothing(plc):
    assert write_plc_row(plc, []) == []
    assert plc.write_calls == []


def test_read_plc_row_single_tag_is_normalised_to_a_row(plc):
    # pycomm3 returns a single Tag (not a list) for a one-tag read; the
    # legacy code iterated that Tag's fields. read_plc_row must still
    # return a one-entry row.
    tag_data, result = read_plc_row(plc, ["FIT_101.Cfg_HiLim"])
    assert tag_data == [("FIT_101.Cfg_HiLim", 90)]
    assert len(result) == 1
    assert all(result)


def test_read_aoi_instances_matches_row_by_row_reads(plc):
    sub_tags = [".Cfg_Desc", ".Cfg_HiLim", ".HiHi.Cfg_Exists"]
    rows = read_aoi_instances(plc, ["FIT_101", "PIT_201"], sub_tags)
    assert len(rows) == 2
    for base, (formatted, raw) in zip(["FIT_101", "PIT_201"], rows):
        expected, expected_raw = read_plc_row(plc, make_tag_list(base, sub_tags))
        assert formatted == expected
        assert [r[0] for r in raw] == [r[0] for r in expected_raw]
        assert all(raw)


def test_read_aoi_instances_chunks_to_the_limit(plc):
    sub_tags = [".Cfg_Desc", ".Cfg_HiLim", ".HiHi.Cfg_Exists"]
    rows = read_aoi_instances(plc, ["FIT_101", "PIT_201"], sub_tags, max_tags_per_read=4)
    assert len(rows) == 2
    # 6 tags in chunks of at most 4: two plc.read calls, 4 then 2 tags.
    assert [len(call) for call in plc.read_log] == [4, 2]
    assert rows[0][0][0] == ("FIT_101.Cfg_Desc", "Flow Transmitter 101")
    assert rows[1][0][2] == ("PIT_201.HiHi.Cfg_Exists", 0)


def test_read_aoi_instances_single_tag_chunks_normalise(plc):
    # max_tags_per_read=1 forces every chunk to a single tag, which
    # pycomm3 returns as one Tag instead of a list; rows must still
    # regroup correctly.
    rows = read_aoi_instances(plc, ["FIT_101"], [".Cfg_Desc", ".Cfg_HiLim"], 1)
    assert rows[0][0] == [
        ("FIT_101.Cfg_Desc", "Flow Transmitter 101"),
        ("FIT_101.Cfg_HiLim", 90),
    ]
    assert len(plc.read_log) == 2


def test_read_aoi_instances_empty_inputs_read_nothing(plc):
    assert read_aoi_instances(plc, [], [".Cfg_HiLim"]) == []
    assert read_aoi_instances(plc, ["FIT_101"], []) == [([], [])]
    assert plc.read_log == []


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
