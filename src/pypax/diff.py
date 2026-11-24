"""Sheet-vs-PLC comparison for write mode.

Replaces the legacy set-based diff, which lost row order and collapsed
duplicate tags. This version preserves spreadsheet row order, keeps each
differing tag once (first occurrence wins), and includes sheet tags the
PLC does not have at all. REAL (float) values compare with a relative
tolerance — they are 32-bit floats in the PLC and rounded to 6 decimals
on read, so a sheet value and its PLC twin can differ by far less than
any meaningful change. INT/BOOL/STRING stay exact.
"""

from __future__ import annotations

import math

# Relative tolerance for REAL (float) comparison. 100.0000001 vs 100.0
# is not a difference; 100.01 vs 100.0 is. Near-zero values still need a
# genuine difference: there is no absolute tolerance by default, so
# 0.0000001 vs 0.0 counts as a difference (and the .6f read rounding
# makes such values read back as 0.0 — see the note in values_equal).
REAL_TOLERANCE = 1e-6

# Absolute tolerance for the write read-back check only. A REAL written
# as 0.0 can read back as a denormal-scale residue (0.0000001) that is
# still a successful write; the plain diff deliberately has no absolute
# tolerance (a sheet value of 0.0000001 is a real edit, not noise).
READBACK_ABSOLUTE_TOLERANCE = 1e-6


def values_equal(sheet_value, plc_value, real_tolerance=REAL_TOLERANCE, absolute_tolerance=0.0):
    """True when two values should count as the same.

    Floats (REALs) compare with relative tolerance (and, when given, an
    absolute tolerance for near-zero values); everything else compares
    exactly (``==``), including bool-vs-int (``True == 1``). ``None`` (a
    missing/unreadable value) on either side is only equal to ``None``.
    """
    if sheet_value is None or plc_value is None:
        return sheet_value is plc_value
    if isinstance(sheet_value, float) or isinstance(plc_value, float):
        return math.isclose(
            sheet_value, plc_value, rel_tol=real_tolerance, abs_tol=absolute_tolerance
        )
    return sheet_value == plc_value


def row_differences(sheet_row, plc_row, real_tolerance=REAL_TOLERANCE, absolute_tolerance=0.0):
    """Differing ``(tag, value)`` entries, in spreadsheet order.

    ``sheet_row`` and ``plc_row`` are lists of ``(tag, value)`` for the
    same instance. An entry differs when the PLC has no such tag or its
    value is not equal (see :func:`values_equal`) to the sheet value.
    """
    plc_values = dict(plc_row)
    seen = set()
    differences = []
    for tag, value in sheet_row:
        if tag in seen:
            continue
        seen.add(tag)
        if tag not in plc_values or not values_equal(
            value, plc_values[tag], real_tolerance, absolute_tolerance
        ):
            differences.append((tag, value))
    return differences
