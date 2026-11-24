"""Spreadsheet (TSV) paste parsing and value validation.

The frontend sends a block copied from a spreadsheet as tab-separated
values. Each line is one instance row; each cell maps to one subtag
column, in order. Values are validated against the column's data type
before anything is staged — an invalid value is rejected with a reason
and never applied.
"""

from __future__ import annotations

import math
import re

from ..string_sizes import merged_sizes

_BOOL_TRUE = {"1", "true", "yes", "on"}
_BOOL_FALSE = {"0", "false", "no", "off"}
_INT_TYPES = ("INT", "DINT", "SINT", "LINT")
_STRING_N_RE = re.compile(r"^STRING_(\d+)$")
# Longest string each type holds, in characters. Pinned sizes live in
# pypax.string_sizes (STRING 82, STRING_Area 8, STRING_Cond /
# STRING_NavTag 20); learned sizes from .L5X imports are merged in
# from config/string_udt_sizes.yaml; STRING_N holds N by the naming
# rule. The merged registry is consulted in _string_limit below.


def _string_limit(data_type_upper: str, registry_path=None) -> int | None:
    """Max characters for a string type, or None when unknown.

    The merged registry (pinned sizes first, then sizes learned from
    ``.L5X`` imports) decides, falling back to the ``STRING_N`` naming
    rule. ``registry_path`` overrides the registry file (tests); the
    pinned map is the fallback when the file is absent.
    """
    merged = merged_sizes(registry_path)
    if data_type_upper in merged:
        return merged[data_type_upper]
    match = _STRING_N_RE.match(data_type_upper)
    return int(match.group(1)) if match else None


def coerce_value(raw: str, data_type: str, registry_path=None):
    """Coerce pasted text to a PLC value for ``data_type``.

    Returns ``(value, None)`` on success or ``(None, reason)`` when the
    text cannot be a value of that type. BOOL becomes ``int`` (0/1),
    matching the core library's convention; STRING types keep the text.

    An empty ``data_type`` means the column's type is unknown — a
    custom AOI whose source ``.L5X`` is gone, so its metadata was
    never cached. Nothing is accepted there: pycomm3 encodes any
    truthy value (including a garbage string) as BOOL true, so an
    unvalidated paste could write a wrong value to the PLC that the
    read-back would only catch afterwards as a mismatch. The user
    recovers by re-importing the ``.L5X``.
    """
    text = raw.strip()
    upper = (data_type or "").upper()
    if not upper:
        return (
            None,
            "column type is unknown (the source .L5X is missing) — "
            "re-import the .L5X to restore type checking",
        )
    if "BOOL" in upper:
        lowered = text.lower()
        if lowered in _BOOL_TRUE:
            return 1, None
        if lowered in _BOOL_FALSE:
            return 0, None
        return None, f"not a valid BOOL (expected 0/1/true/false), got {raw!r}"
    if any(kind in upper for kind in _INT_TYPES):
        try:
            return int(text), None
        except ValueError:
            return None, f"not a valid {data_type}, got {raw!r}"
    if "REAL" in upper:
        try:
            value = float(text)
        except ValueError:
            return None, f"not a valid REAL, got {raw!r}"
        if not math.isfinite(value):
            return None, f"not a valid REAL, got {raw!r}"
        return value, None
    if "STRING" in upper:
        # The spreadsheet flags a string longer than its type holds;
        # the PLC would not hold it either (pycomm3's string encoder
        # does no truncation), so reject it here like any other bad
        # value instead of staging a write that cannot land as-is.
        limit = _string_limit(upper, registry_path)
        if limit is not None and len(raw) > limit:
            return None, (
                f"string is {len(raw)} characters, longer than the "
                f"{limit}-character limit for {data_type}"
            )
        return raw, None
    return raw, None


def parse_tsv(
    tsv: str,
    instances: list[str],
    columns,
    anchor_instance: str | None = None,
    anchor_subtag: str | None = None,
    registry_path=None,
) -> tuple[list[dict], list[dict]]:
    """Parse pasted TSV into staged changes and rejections.

    ``columns`` are :class:`Column` objects (subtag + data type) in grid
    order. The pasted block's top-left cell maps to the anchor
    (``anchor_instance``, ``anchor_subtag``); cells then flow right
    across subtags in column order and down across instances in row
    order. Without an anchor, the block starts at the first instance
    and first subtag, except that a line whose first cell names a known
    instance pastes into that instance (the rest of the line is values).
    Empty cells are skipped (no change). Leading/trailing blank lines
    are ignored; a blank line inside the block is an empty row that
    still advances the row position, so later rows keep their instance
    mapping. Rejections are
    ``{"row", "column", "reason"}`` where ``row`` is the instance name
    the cell belongs to (the pasted first cell when a line maps to no
    instance at all); changes are
    ``{"instance", "subtag", "tag", "value"}`` in cell order.
    """
    changes: list[dict] = []
    rejected: list[dict] = []
    subtag_names = [column.subtag for column in columns]
    anchored = anchor_instance is not None or anchor_subtag is not None

    if anchor_instance is not None and anchor_instance not in instances:
        rejected.append(
            {
                "row": anchor_instance,
                "column": anchor_subtag or "",
                "reason": f"unknown instance {anchor_instance!r}",
            }
        )
        return changes, rejected
    if anchor_subtag is not None and anchor_subtag not in subtag_names:
        rejected.append(
            {
                "row": anchor_instance
                if anchor_instance is not None
                else (instances[0] if instances else ""),
                "column": anchor_subtag,
                "reason": f"unknown subtag {anchor_subtag!r}",
            }
        )
        return changes, rejected

    start_instance = instances.index(anchor_instance) if anchor_instance is not None else 0
    start_column = subtag_names.index(anchor_subtag) if anchor_subtag is not None else 0

    lines = tsv.splitlines()
    # Leading/trailing blank lines are not part of the pasted block.
    # A blank line *inside* the block is a real (empty) row: it still
    # advances the row position, so later rows keep their instance
    # mapping. Skipping it instead would silently land the next row's
    # values on the wrong instance.
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()

    data_line = 0  # 0-based index over block lines, for positional mapping
    for line in lines:
        if not line.strip():
            data_line += 1  # empty row: no changes, but the position counts
            continue
        cells = [cell.strip() for cell in line.split("\t")]
        if anchored:
            instance_index = start_instance + data_line
            if instance_index >= len(instances):
                rejected.append(
                    {
                        "row": cells[0],
                        "column": "",
                        "reason": f"no instance for pasted row {data_line + 1}",
                    }
                )
                data_line += 1
                continue
            instance = instances[instance_index]
            value_cells = cells
            first_column = start_column
        elif cells[0] in instances:
            instance = cells[0]
            value_cells = cells[1:]
            first_column = 0
        else:
            if data_line >= len(instances):
                rejected.append(
                    {
                        "row": cells[0],
                        "column": "",
                        "reason": f"no instance for pasted row {data_line + 1}",
                    }
                )
                data_line += 1
                continue
            instance = instances[data_line]
            value_cells = cells
            first_column = 0
        data_line += 1
        for offset, raw in enumerate(value_cells):
            column_index = first_column + offset
            if column_index >= len(columns):
                break  # extra cells past the last column are ignored
            if raw == "":
                continue
            column = columns[column_index]
            value, reason = coerce_value(raw, column.data_type, registry_path)
            if reason is not None:
                rejected.append({"row": instance, "column": column.subtag, "reason": reason})
            else:
                changes.append(
                    {
                        "instance": instance,
                        "subtag": column.subtag,
                        "tag": instance + column.subtag,
                        "value": value,
                    }
                )
    return changes, rejected
