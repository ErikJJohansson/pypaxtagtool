"""AOI subtag definitions as versioned data files.

The runtime source is still ``AOI_definitions.py``; this module reads
and writes the generated ``config/aoi/vX.Y.Z.yaml`` record of the same
data (one AOI name mapping to its ordered list of subtags), produced
from the template spreadsheet by ``scripts/extract_aoi_definitions.py``.

Only the generated shape is supported here: a top-level mapping of AOI
name to a list of subtag strings. The loader is a real YAML parser
(``yaml.safe_load``) with two strictness rules the generated files rely
on: duplicate AOI keys are rejected, and anything outside the expected
shape raises ``ValueError`` instead of silently misreading it.
"""

from __future__ import annotations

from pathlib import Path

import yaml

# A subtag must fit on one single-quoted YAML line. PlantPAX subtags
# are dotted member paths (``.Cfg_HiLim``, ``.Cfg_NavTag[0]``); anything
# with a newline or control character is rejected at dump time.
_SAFE_SUBTAG_CHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._[]")


class _StrictLoader(yaml.SafeLoader):
    """SafeLoader that rejects duplicate mapping keys."""


def _construct_mapping(loader, node):
    mapping = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=True)
        if key in mapping:
            raise ValueError(f"duplicate AOI key: {key!r}")
        mapping[key] = loader.construct_object(value_node, deep=True)
    return mapping


_StrictLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, _construct_mapping)


def dump_aoi_definitions(definitions, path):
    """Write ``{AOI: [subtag, ...]}`` as YAML to ``path``.

    AOIs and subtags are written in the given (insertion) order, each
    subtag single-quoted. Returns the path written.
    """
    path = Path(path)
    lines = [
        "# AOI subtag definitions for the PlantPAX v4.10.6 library.",
        "# Generated from ProcessLibraryOnlineConfigTool.xlsm by",
        "# scripts/extract_aoi_definitions.py -- do not edit by hand.",
        "# Definitions carried from v4.10.4 (logic-only changes).",
    ]
    for aoi, subtags in definitions.items():
        if ":" in aoi or "\n" in aoi:
            raise ValueError(f"AOI name not safe for YAML output: {aoi!r}")
        lines.append(f"{aoi}:")
        for subtag in subtags:
            if not set(subtag) <= _SAFE_SUBTAG_CHARS:
                raise ValueError(f"subtag not safe for YAML output: {subtag!r}")
            lines.append(f"  - '{subtag}'")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


def load_aoi_definitions(path):
    """Read a generated AOI YAML file into ``{AOI: [subtag, ...]}``.

    Strict about shape: the document must be a mapping of AOI name to
    a list of subtag strings, with no duplicate AOI keys. Anything else
    raises ``ValueError``.
    """
    path = Path(path)
    try:
        data = yaml.load(path.read_text(encoding="utf-8"), Loader=_StrictLoader)
    except yaml.YAMLError as exc:
        raise ValueError(f"{path}: invalid YAML: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"{path}: expected a mapping of AOI to subtag list")
    definitions = {}
    for aoi, subtags in data.items():
        if not isinstance(aoi, str) or not aoi:
            raise ValueError(f"{path}: AOI name must be a non-empty string: {aoi!r}")
        if not isinstance(subtags, list) or not all(isinstance(s, str) for s in subtags):
            raise ValueError(f"{path}: {aoi}: expected a list of subtag strings")
        definitions[aoi] = list(subtags)
    return definitions
