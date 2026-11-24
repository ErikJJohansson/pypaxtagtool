"""String-UDT size registry: pinned sizes + sizes learned from ``.L5X`` imports.

Pinned sizes are ground truth from exports Erik supplied and are
never overridden: the Logix ``STRING`` holds 82 data characters,
``STRING_Area`` holds 8 (P_ValveMO 4.10.06 export), ``STRING_Cond``
and ``STRING_NavTag`` hold 20 each (P_Intlk 4.10.06 export).
``STRING_N`` holds N by the naming rule.

Learned sizes come from :func:`pypax.l5x.learn_string_sizes` at
import-confirm time and persist in ``config/string_udt_sizes.yaml``
as ``{NAME: {size: <int>, source: <filename>}}`` — type names, sizes,
and source filenames only, never PLC or site details. A learned size
that conflicts with a pinned size is ignored (pinned wins) and
reported as a warning; a learned size that conflicts with an earlier
learned size for the same type is likewise kept as-is and warned
about — sizes are never guessed and never silently overwritten.
"""

from __future__ import annotations

import re
from functools import lru_cache
from pathlib import Path

import yaml

# Pinned sizes: verified against real exports, never overridden.
PINNED_SIZES = {"STRING": 82, "STRING_AREA": 8, "STRING_COND": 20, "STRING_NAVTAG": 20}

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_REGISTRY_PATH = REPO_ROOT / "config" / "string_udt_sizes.yaml"

_TYPE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
_STRING_N_RE = re.compile(r"^STRING_(\d+)$")


def _resolve_registry_path(path: str | Path | None) -> Path:
    return Path(path) if path is not None else DEFAULT_REGISTRY_PATH


def _valid_entry(name: object, entry: object) -> tuple[str, int, str] | None:
    """Validate one registry entry; ``None`` when it is not usable.

    A usable entry is ``{NAME: {size: positive int, source: str}}``
    with a well-formed type name. Anything else is skipped, never
    guessed — a hand-edited registry file must not break validation.
    """
    if not isinstance(name, str) or not _TYPE_NAME_RE.match(name):
        return None
    if not isinstance(entry, dict):
        return None
    size = entry.get("size")
    source = entry.get("source")
    if not isinstance(size, int) or isinstance(size, bool) or size <= 0:
        return None
    if not isinstance(source, str) or not source:
        return None
    return name.upper(), size, source


@lru_cache(maxsize=16)
def _load_registry_cached(path_str: str) -> dict[str, dict[str, object]]:
    """Learned entries from the registry file, keyed by upper-cased name.

    A missing or unreadable file means no learned sizes (the pinned
    map is the fallback); malformed entries are skipped. Cached per
    path — :func:`merge_learned_sizes` clears the cache after writing.
    """
    path = Path(path_str)
    try:
        text = path.read_text(encoding="utf-8")
    except OSError:
        return {}
    try:
        data = yaml.safe_load(text)
    except yaml.YAMLError:
        return {}
    if not isinstance(data, dict):
        return {}
    entries: dict[str, dict[str, object]] = {}
    for name, entry in data.items():
        valid = _valid_entry(name, entry)
        if valid is None:
            continue
        upper, size, source = valid
        entries[upper] = {"size": size, "source": source}
    return entries


def load_learned_sizes(path: str | Path | None = None) -> dict[str, int]:
    """Learned sizes ``{NAME: size}`` from the registry file ({} when absent)."""
    resolved = _resolve_registry_path(path)
    return {
        name: int(entry["size"]) for name, entry in _load_registry_cached(str(resolved)).items()
    }


def merged_sizes(path: str | Path | None = None) -> dict[str, int]:
    """Pinned sizes overlaid on learned sizes (pinned wins on conflict)."""
    return {**load_learned_sizes(path), **PINNED_SIZES}


def merge_learned_sizes(
    new: dict[str, int], source: str, path: str | Path | None = None
) -> list[str]:
    """Persist ``new`` learned sizes into the registry file.

    ``source`` is the basename of the ``.L5X`` the sizes were learned
    from. Returns warning strings for every size that was *not*
    stored: a size conflicting with a pinned size, or with an earlier
    learned size for the same type (first-seen wins). Unchanged sizes
    are stored silently; the file is only rewritten when something
    actually changed.
    """
    resolved = _resolve_registry_path(path)
    entries = dict(_load_registry_cached(str(resolved)))
    warnings: list[str] = []
    changed = False
    for raw_name, size in sorted(new.items()):
        name = raw_name.upper()
        match = _STRING_N_RE.match(name)
        if match and int(match.group(1)) != size:
            warnings.append(
                f"learned size {size} for {name} from {source} conflicts "
                f"with the STRING_N naming rule ({name} holds "
                f"{match.group(1)}) — ignored"
            )
            continue
        if name in PINNED_SIZES:
            if PINNED_SIZES[name] != size:
                warnings.append(
                    f"learned size {size} for {name} conflicts with the pinned "
                    f"size {PINNED_SIZES[name]} from {source} — ignored"
                )
            continue
        existing = entries.get(name)
        if existing is not None:
            if existing["size"] != size:
                warnings.append(
                    f"learned size {size} for {name} from {source} conflicts "
                    f"with the earlier learned size {existing['size']} "
                    f"(from {existing['source']}) — kept the earlier size"
                )
            continue
        entries[name] = {"size": size, "source": source}
        changed = True
    if changed:
        resolved.parent.mkdir(parents=True, exist_ok=True)
        text = yaml.safe_dump(
            {name: entries[name] for name in sorted(entries)},
            sort_keys=True,
            allow_unicode=False,
        )
        resolved.write_text(text, encoding="utf-8")
        _load_registry_cached.cache_clear()
    return warnings
