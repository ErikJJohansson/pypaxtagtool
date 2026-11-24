"""PLC access helpers: tag discovery, bulk row read/write, failures.

Moved verbatim in behaviour from the legacy ``pypax.py`` CLI, including
the exact type coercion in :func:`read_plc_row` (BOOL to ``int``, REAL
with ``.6f`` / ``.6e`` formatting, everything else passed through).
"""

from __future__ import annotations

from itertools import product


def get_aoi_tag_instances(plc, tag_type):
    """Base tag names of all non-alias instances of an AOI type.

    Array instances are expanded to one entry per element
    (``TAG[0]``, ``TAG[1]``, ...).
    """
    tag_list = []
    for tag, _def in plc.tags.items():
        if _def["data_type_name"] == tag_type and not (_def["alias"]):
            if _def["dim"] > 0:
                tag_list = tag_list + get_dim_list(tag, _def["dimensions"])
            else:
                tag_list.append(tag)
    return tag_list


def get_dim_list(base_tag, dim_list):
    """Expand array dimensions into one tag name per element.

    Zero dimensions are dropped; ``[2, 3]`` yields ``TAG[0][0]`` ...
    ``TAG[1][2]`` in row-major (``itertools.product``) order.
    """
    filtered_list = list(filter(lambda num: num != 0, dim_list))
    temp = []
    for indices in product(*[range(dim) for dim in filtered_list]):
        temp.append(base_tag + "".join(f"[{i}]" for i in indices))
    return temp


def make_tag_list(base_tag, sub_tags):
    """Full tag paths for one instance: ``base_tag + sub_tag`` each."""
    return [base_tag + s for s in sub_tags]


def read_plc_row(plc, tag_list):
    """Read one instance row; returns ``(formatted, raw_results)``.

    Formatted entries are ``(tag_name, value)`` with BOOL coerced to
    ``int`` and REAL rounded to 6 decimals (scientific form for very
    small/large values, ``int`` for whole-number REALs).
    Raw results are normalised to a list first: pycomm3 returns a single
    ``Tag`` (not a list) when exactly one tag is read, and iterating that
    would walk the tag's fields instead of its rows. The legacy code
    never hit this (every AOI row reads 2+ subtags), but a single-tag
    caller would have read garbage.
    """
    raw = plc.read(*tag_list)
    tag_data = raw if isinstance(raw, list) else [raw]

    tag_data_formatted = []
    for s in tag_data:
        if s[2] == "BOOL":
            data = int(s[1])
        elif s[2] == "REAL":
            if "e" in str(s[1]):
                data = float(format(s[1], ".6e"))
            elif s[1].is_integer():
                data = int(s[1])
            else:
                data = float(format(s[1], ".6f"))
        else:
            data = s[1]
        tag_data_formatted.append((s[0], data))

    return tag_data_formatted, tag_data


def write_plc_row(plc, tag_data):
    """Write ``(tag, value)`` pairs; always returns a list of bools."""
    result = plc.write(*tag_data)
    if len(tag_data) == 1:
        return [result]
    else:
        return result


def get_failed_tags(tag_list, result_list):
    """Tag names whose corresponding result is falsy, in order."""
    failed_indexes = [i for i, val in enumerate(result_list) if not val]
    return [tag_list[i] for i in failed_indexes]


def failed_tag_formatter(tags, write_mode):
    """Human-readable warning block listing tags that failed."""
    if write_mode:
        return "****WARNING****\nCANNOT WRITE TO\n" + ", ".join(tags) + "\n***************"
    else:
        return "****WARNING****\nCANNOT READ FROM\n" + ", ".join(tags) + "\n***************"
