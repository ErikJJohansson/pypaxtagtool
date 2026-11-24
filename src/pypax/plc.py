"""PLC access helpers: tag discovery, bulk row read/write, failures.

Moved verbatim in behaviour from the legacy ``pypax.py`` CLI, including
the exact type coercion in :func:`read_plc_row` (BOOL to ``int``, REAL
with ``.6f`` / ``.6e`` formatting, everything else passed through).
"""

from __future__ import annotations

from itertools import product

# pycomm3 sends one CIP multi-request packet per read()/write() call.
# Requests that are too large are rejected by the PLC, so bulk work is
# split into chunks of at most this many tags per call. 100 keeps a
# chunk comfortably inside one packet for typical PlantPAX config tags
# (short STRINGs, REALs, DINTs, BOOLs).
MAX_TAGS_PER_READ = 100
MAX_TAGS_PER_WRITE = 100


def _is_base_tag_name(name):
    """True if a tags-dict key names a base instance, not a member path.

    Real ``LogixDriver.tags`` keys are base instances (``FIT_101``) or
    program-scoped bases (``Program:MainProgram.FIT_101``); member paths
    such as ``FIT_101.Cfg_HiLim`` never appear there. The FakePLC seeds
    member entries too (its tags dict doubles as the value store), so
    discovery must ignore them: a name is a base instance when, after
    stripping an optional ``Program:<name>.`` prefix, it contains no ``.``.
    ControlLogix base tag names cannot contain a period, so on a real PLC
    this filter changes nothing.
    """
    if name.startswith("Program:"):
        _, _, name = name.partition(".")
    return "." not in name


def build_tag_index(plc):
    """Group base-instance tags by data type, in first-seen order.

    One pass over ``plc.tags``; aliases and member paths are skipped
    (see :func:`_is_base_tag_name`). Returns
    ``{data_type_name: [(tag, definition), ...]}``. Callers that scan
    several AOI types should build the index once and pass it to
    :func:`get_aoi_tag_instances` instead of rescanning ``plc.tags``
    per AOI.
    """
    index = {}
    for tag, _def in plc.tags.items():
        if _def["alias"] or not _is_base_tag_name(tag):
            continue
        index.setdefault(_def["data_type_name"], []).append((tag, _def))
    return index


def _get_tag_index(plc, index=None):
    """An explicit index, or the index cached on the plc object.

    The cache is keyed by the tags dict's identity and size, so replacing
    ``plc.tags`` or adding/removing entries rebuilds it; in-place edits
    to a single entry's fields are not detected (``plc.tags`` is built
    once at connect time and not mutated afterwards on a real PLC).
    Objects that refuse new attributes simply rebuild on every call.
    """
    if index is not None:
        return index
    fingerprint = (id(plc.tags), len(plc.tags))
    cached = getattr(plc, "_pypax_tag_index", None)
    if cached is not None and cached[0] == fingerprint:
        return cached[1]
    fresh = build_tag_index(plc)
    try:
        plc._pypax_tag_index = (fingerprint, fresh)
    except Exception:
        pass
    return fresh


def get_aoi_tag_instances(plc, tag_type, index=None):
    """Base tag names of all non-alias instances of an AOI type.

    Array instances are expanded to one entry per element
    (``TAG[0]``, ``TAG[1]``, ...). Pass a prebuilt ``index`` from
    :func:`build_tag_index` to avoid rescanning ``plc.tags``; otherwise
    the index is built once and cached on the plc object.
    """
    tag_list = []
    for tag, _def in _get_tag_index(plc, index).get(tag_type, []):
        if _def["dim"] > 0:
            tag_list.extend(get_dim_list(tag, _def["dimensions"]))
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


def _format_tag_results(tag_data):
    """Format raw Tag results to ``(tag_name, value)`` pairs.

    BOOL coerces to ``int`` and REAL rounds to 6 decimals (scientific
    form for very small/large values, ``int`` for whole-number REALs).
    """
    tag_data_formatted = []
    for s in tag_data:
        if s[1] is None:
            # Failed read (value None, whatever the reported type): the
            # raw result is already falsy, so report the tag as failed
            # instead of crashing on int(None) / formatting None as REAL.
            tag_data_formatted.append((s[0], None))
            continue
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
    return tag_data_formatted


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
    return _format_tag_results(tag_data), tag_data


def read_aoi_instances(plc, base_tags, sub_tags, max_tags_per_read=MAX_TAGS_PER_READ):
    """Read all instances of one AOI in chunked bulk reads.

    Returns one ``(formatted_row, raw_results)`` pair per base tag, in
    ``base_tags`` order, each row in ``sub_tags`` order — the same shape
    :func:`read_plc_row` returns for a single instance. All
    (instance x subtag) tags are flattened and read in chunks of at most
    ``max_tags_per_read`` so no single request exceeds the CIP
    multi-request packet limit; the legacy one-read-per-row loop issued
    ``len(base_tags)`` requests of up to ``len(sub_tags)`` tags each.
    """
    all_tags = [base + s for base in base_tags for s in sub_tags]
    per_row = len(sub_tags)
    flat_formatted = []
    flat_raw = []
    for start in range(0, len(all_tags), max_tags_per_read):
        chunk = all_tags[start : start + max_tags_per_read]
        raw = plc.read(*chunk)
        chunk_results = raw if isinstance(raw, list) else [raw]
        flat_raw.extend(chunk_results)
        flat_formatted.extend(_format_tag_results(chunk_results))
    return [
        (
            flat_formatted[i * per_row : (i + 1) * per_row],
            flat_raw[i * per_row : (i + 1) * per_row],
        )
        for i in range(len(base_tags))
    ]


def write_plc_row(plc, tag_data, max_tags_per_write=MAX_TAGS_PER_WRITE):
    """Write ``(tag, value)`` pairs; always returns a list of results.

    Pairs are written in sequential chunks of at most
    ``max_tags_per_write`` so no single request exceeds the CIP
    multi-request packet limit. Results are combined in original order;
    a single-Tag result from pycomm3 (a one-pair chunk) is normalised
    to a one-entry list. No pairs means no PLC call and an empty list.
    """
    if not tag_data:
        return []
    results = []
    for start in range(0, len(tag_data), max_tags_per_write):
        chunk = tag_data[start : start + max_tags_per_write]
        result = plc.write(*chunk)
        results.extend(result if isinstance(result, list) else [result])
    return results


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
