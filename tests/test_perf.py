"""Call-count (perf) tests for the Phase 2 PLC optimizations.

These assert request counts against the FakePLC's call logs, not wall
time: the optimizations are about staying inside CIP multi-request
packet limits and avoiding one request per row, both of which are
exactly countable. No live PLC is involved.
"""

from pypax.plc import (
    MAX_TAGS_PER_READ,
    MAX_TAGS_PER_WRITE,
    read_aoi_instances,
    write_plc_row,
)
from fakes.fake_plc import FakePLC

SUBTAGS_101 = [f".Cfg_Sub{i:03d}" for i in range(101)]


def _seed_instances(plc, count, sub_tags, base_prefix="TT"):
    """Seed ``count`` P_AIn instances; value of (i, j) is i * 1000 + j."""
    bases = []
    for i in range(count):
        base = f"{base_prefix}_{i:03d}"
        bases.append(base)
        plc.add_instance(
            base,
            "P_AIn",
            [(sub, "REAL", float(i * 1000 + j)) for j, sub in enumerate(sub_tags)],
        )
    return bases


def test_bulk_read_50_instances_101_subtags_uses_bounded_calls():
    plc = FakePLC()
    plc.open()
    try:
        bases = _seed_instances(plc, 50, SUBTAGS_101)
        # A per-row read would issue 50 requests of 101 tags each — every
        # one over the 100-tag packet limit. The bulk read must stay
        # inside it: ceil(50 * 101 / 100) = 51 calls, none over 100 tags.
        assert len(SUBTAGS_101) > MAX_TAGS_PER_READ
        rows = read_aoi_instances(plc, bases, SUBTAGS_101)
        assert len(rows) == 50
        assert len(plc.read_log) == 51
        assert max(len(call) for call in plc.read_log) <= MAX_TAGS_PER_READ
        assert sum(len(call) for call in plc.read_log) == 50 * 101
        for i, (formatted, raw) in enumerate(rows):
            assert len(formatted) == 101
            assert formatted[0] == (f"TT_{i:03d}.Cfg_Sub000", i * 1000)
            assert formatted[100] == (f"TT_{i:03d}.Cfg_Sub100", i * 1000 + 100)
            assert all(raw)
    finally:
        plc.close()


def test_bulk_read_many_small_rows_uses_far_fewer_calls_than_per_row():
    plc = FakePLC()
    plc.open()
    try:
        sub_tags = [".Cfg_HiLim", ".Cfg_LoLim"]
        bases = _seed_instances(plc, 100, sub_tags)
        rows = read_aoi_instances(plc, bases, sub_tags)
        assert len(rows) == 100
        # 200 tags in 100-tag chunks: 2 bulk calls vs 100 one-per-row.
        assert len(plc.read_log) == 2
        assert rows[99][0] == [
            ("TT_099.Cfg_HiLim", 99000),
            ("TT_099.Cfg_LoLim", 99001),
        ]
    finally:
        plc.close()


def test_bulk_write_250_tags_uses_three_calls():
    plc = FakePLC()
    plc.open()
    try:
        bases = _seed_instances(plc, 50, SUBTAGS_101)
        tag_data = [(f"{base}{sub}", 1.5) for base in bases[:2] for sub in SUBTAGS_101] + [
            (f"{bases[2]}{sub}", 2.5) for sub in SUBTAGS_101[:48]
        ]
        assert len(tag_data) == 250
        results = write_plc_row(plc, tag_data)
        assert len(results) == 250
        assert all(results)
        # 250 pairs in chunks of at most 100: three plc.write calls.
        assert [len(call) for call in plc.write_calls] == [100, 100, 50]
        assert len(plc.write_calls) == (250 + MAX_TAGS_PER_WRITE - 1) // MAX_TAGS_PER_WRITE
        assert plc.read("TT_000.Cfg_Sub000").value == 1.5
        assert plc.read("TT_002.Cfg_Sub047").value == 2.5
    finally:
        plc.close()
