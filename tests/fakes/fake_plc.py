"""In-memory fake PLC for tests.

A stand-in for pycomm3's ``LogixDriver`` implementing the exact subset of
its interface that ``pypax.py`` (and later the webapp core) relies on:

- ``open()`` / ``close()``
- ``get_plc_name()``
- ``tags`` — a dict of ``{name: {...}}`` with the keys the CLI reads:
  ``data_type_name``, ``dim``, ``dimensions``, ``alias``
- ``read(*tag_names)`` / ``write(*tag_value_pairs)`` — return one
  ``FakeTag`` when a single tag is requested and a list otherwise,
  exactly like pycomm3's ``LogixDriver``; ``FakeTag`` is a 4-field
  tuple ``(tag, value, type, error)`` like pycomm3's ``Tag``

No network access, no real PLC, no pycomm3 import. Pure Python, so it runs
identically on macOS, Windows, and Linux.

Seeded with two realistic ``P_AIn`` (analog input) instances plus one
``P_DIn`` instance, so AOI discovery by ``data_type_name`` can be tested.
"""

from __future__ import annotations


class FakeTag:
    """One tag read/write result, mimicking pycomm3's ``Tag`` exactly.

    pycomm3 1.2.16 (the pinned version) defines ``Tag`` as a NamedTuple
    with fields ``(tag, value, type, error)`` — 4 fields, indexable and
    unpackable as such. ``bool(tag)`` is True only when the value is not
    None and there is no error. ``.error`` is both the 4th tuple field
    and an attribute, exactly like the real thing.
    """

    __slots__ = ("tag", "value", "type", "error")

    def __init__(self, tag, value, type_, error=None):
        self.tag = tag
        self.value = value
        self.type = type_
        self.error = error

    def __getitem__(self, index):
        return (self.tag, self.value, self.type, self.error)[index]

    def __iter__(self):
        yield self.tag
        yield self.value
        yield self.type
        yield self.error

    def __len__(self):
        return 4

    def __bool__(self):
        return self.value is not None and self.error is None

    def __repr__(self):
        return (
            f"FakeTag(tag={self.tag!r}, value={self.value!r}, "
            f"type={self.type!r}, error={self.error!r})"
        )


# (tag name, data type, seed value)
_SEED_TAGS = [
    # P_AIn instance 1: flow transmitter
    ("FIT_101.Cfg_Desc", "STRING", "Flow Transmitter 101"),
    ("FIT_101.Cfg_Label", "STRING", "FIT-101"),
    ("FIT_101.Cfg_Tag", "STRING", "FIT-101"),
    ("FIT_101.Cfg_EU", "STRING", "m3/h"),
    ("FIT_101.Cfg_HiLim", "REAL", 90.0),
    ("FIT_101.Cfg_LoLim", "REAL", 10.0),
    ("FIT_101.Cfg_HiHiLim", "REAL", 95.0),
    ("FIT_101.Cfg_LoLoLim", "REAL", 5.0),
    ("FIT_101.HiHi.Cfg_Exists", "BOOL", True),
    ("FIT_101.HiHi.Cfg_Severity", "DINT", 500),
    # P_AIn instance 2: pressure transmitter
    ("PIT_201.Cfg_Desc", "STRING", "Pressure Transmitter 201"),
    ("PIT_201.Cfg_Label", "STRING", "PIT-201"),
    ("PIT_201.Cfg_Tag", "STRING", "PIT-201"),
    ("PIT_201.Cfg_EU", "STRING", "bar"),
    ("PIT_201.Cfg_HiLim", "REAL", 8.0),
    ("PIT_201.Cfg_LoLim", "REAL", 0.5),
    ("PIT_201.Cfg_HiHiLim", "REAL", 9.0),
    ("PIT_201.Cfg_LoLoLim", "REAL", 0.0),
    ("PIT_201.HiHi.Cfg_Exists", "BOOL", False),
    ("PIT_201.HiHi.Cfg_Severity", "DINT", 750),
    # P_DIn instance: a discrete input, so AOI discovery can be filtered
    ("XIC_301.Cfg_Desc", "STRING", "Pump Running Feedback"),
    ("XIC_301.Cfg_Label", "STRING", "XIC-301"),
    ("XIC_301.Cfg_Debounce", "DINT", 100),
]


def _data_type_name(tag_name):
    """AOI type of a tag, from its base instance name."""
    base = tag_name.split(".", 1)[0]
    if base.startswith("FIT_") or base.startswith("PIT_"):
        return "P_AIn"
    if base.startswith("XIC_"):
        return "P_DIn"
    return "UNKNOWN"


class FakePLC:
    """In-memory LogixDriver stand-in. No network, no hardware."""

    def __init__(self, commpath="10.0.0.1/0", plc_name="TestPLC"):
        self.commpath = commpath
        self._plc_name = plc_name
        self._open = False
        self._values = {name: value for name, _type, value in _SEED_TAGS}
        self._types = {name: type_ for name, type_, _value in _SEED_TAGS}
        self.tags = {
            name: {
                "data_type_name": _data_type_name(name),
                "dim": 0,
                "dimensions": [0, 0, 0],
                "alias": False,
            }
            for name in self._values
        }
        # Every write, in order, as (tag, value) — lets tests assert
        # batching/chunking behaviour without a real PLC.
        self.write_log = []

    def open(self):
        self._open = True
        return True

    def close(self):
        self._open = False

    def get_plc_name(self):
        if not self._open:
            raise RuntimeError("PLC connection is not open")
        return self._plc_name

    def _require_open(self):
        if not self._open:
            raise RuntimeError("PLC connection is not open")

    def read(self, *tag_names):
        """Read tags, mirroring pycomm3: one tag -> one FakeTag, many -> list."""
        self._require_open()
        results = []
        for name in tag_names:
            if name in self._values:
                results.append(FakeTag(name, self._values[name], self._types[name]))
            else:
                results.append(FakeTag(name, None, None, error="Unknown tag"))
        if len(results) == 1:
            return results[0]
        return results

    def write(self, *tag_value_pairs):
        """Write (tag, value) pairs, mirroring pycomm3.

        Returns one FakeTag per pair (a single FakeTag when only one pair
        was written). A written tag echoes the written value with no error;
        an unknown tag yields value None and an error, so it is falsy.
        """
        self._require_open()
        results = []
        for tag, value in tag_value_pairs:
            if tag in self._values:
                self._values[tag] = value
                self.write_log.append((tag, value))
                results.append(FakeTag(tag, value, self._types[tag]))
            else:
                results.append(FakeTag(tag, value, None, error="Unknown tag"))
        if len(results) == 1:
            return results[0]
        return results

    def reset(self):
        """Restore seed values and clear the write log."""
        self._values = {name: value for name, _type, value in _SEED_TAGS}
        self.write_log.clear()
