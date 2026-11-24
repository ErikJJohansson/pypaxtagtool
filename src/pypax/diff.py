"""Sheet-vs-PLC comparison for write mode.

Replaces the legacy set-based diff, which lost row order and collapsed
duplicate tags. This version preserves spreadsheet row order, keeps each
differing tag once (first occurrence wins), and includes sheet tags the
PLC does not have at all. Comparison is exact for now; tolerance-aware
comparison for REALs is Phase 2 work.
"""

from __future__ import annotations


def row_differences(sheet_row, plc_row):
    """Differing ``(tag, value)`` entries, in spreadsheet order.

    ``sheet_row`` and ``plc_row`` are lists of ``(tag, value)`` for the
    same instance. An entry differs when the PLC has no such tag or its
    value is not equal (``==``) to the sheet value.
    """
    plc_values = dict(plc_row)
    seen = set()
    differences = []
    for tag, value in sheet_row:
        if tag in seen:
            continue
        seen.add(tag)
        if tag not in plc_values or plc_values[tag] != value:
            differences.append((tag, value))
    return differences
