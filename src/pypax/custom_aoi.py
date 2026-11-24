"""Custom AOI templates imported from Studio 5000 ``.L5X`` exports.

A confirmed import is stored as one YAML file per AOI under
``config/aoi/custom/<AOI>.yaml``. The file has the same shape as the
generated ``config/aoi/v4.10.6.yaml`` (one AOI name mapping to its
ordered list of subtag strings) so the existing
:func:`pypax.aoi.load_aoi_definitions` loader reads it unchanged, plus
a comment header recording the source export, its TargetRevision, and
the import date. Re-importing an AOI overwrites its own file; a name
that collides with a PlantPAX library AOI in the template spreadsheet
is rejected by :func:`check_collision` (call it at proposal time —
``save_custom_aoi`` enforces it too when writing to the default
location).

Nothing here touches the PLC or the template workbook: the spreadsheet
is only ever opened read-only to list its AOI sheets.
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import date
from pathlib import Path

import openpyxl

from .aoi import load_aoi_definitions
from .layout import get_aoi_list
from .l5x import Column

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TEMPLATE_PATH = REPO_ROOT / "ProcessLibraryOnlineConfigTool.xlsm"
DEFAULT_CUSTOM_DIR = REPO_ROOT / "config" / "aoi" / "custom"

_AOI_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
# Same single-quote safety rule as pypax.aoi.dump_aoi_definitions: a
# subtag must fit on one single-quoted YAML line.
_SAFE_SUBTAG_CHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789._[]")
# The header line recording the SHA-256 of the source .L5X at confirm
# time; the folder scan compares it against the file on disk to spot a
# changed export. Written and read by save_custom_aoi /
# read_source_sha256 below. Hex only, so it is header-safe by
# construction.
_SHA256_HEADER_RE = re.compile(r"^# Source SHA-256: ([0-9a-f]{64})$", re.MULTILINE)


def _header_safe(value: str) -> str:
    """One line of a YAML comment header.

    A newline in the source file name or revision would break out of
    the ``#`` comment and inject YAML lines into the custom-AOI file —
    e.g. a smuggled second AOI definition that bypasses name
    validation and collision checks entirely. Collapse all line
    breaks to spaces.
    """
    return " ".join(str(value).splitlines())


def _validate_aoi_name(name: str) -> None:
    if not isinstance(name, str) or not _AOI_NAME_RE.match(name):
        raise ValueError(
            f"AOI name must be a valid identifier (letters, digits, "
            f"underscore, not starting with a digit): {name!r}"
        )


def _subtags_of(columns) -> list[str]:
    subtags: list[str] = []
    for column in columns:
        subtag = column if isinstance(column, str) else column.subtag
        if not isinstance(subtag, str) or not subtag:
            raise ValueError(f"column subtag must be a non-empty string: {subtag!r}")
        if not set(subtag) <= _SAFE_SUBTAG_CHARS:
            raise ValueError(f"subtag not safe for YAML output: {subtag!r}")
        subtags.append(subtag)
    if len(set(subtags)) != len(subtags):
        raise ValueError("duplicate subtags in custom AOI columns")
    return subtags


def _template_aoi_names(template_path: str | Path | None) -> set[str]:
    if template_path is None:
        return set()
    template_path = Path(template_path)
    if not template_path.exists():
        return set()
    book = openpyxl.load_workbook(template_path, read_only=True, data_only=True)
    try:
        return set(get_aoi_list(book))
    finally:
        book.close()


def check_collision(
    name: str,
    template_path: str | Path | None = DEFAULT_TEMPLATE_PATH,
    custom_dir: str | Path | None = DEFAULT_CUSTOM_DIR,
) -> None:
    """Reject ``name`` if it already names a PlantPAX or custom AOI.

    Raises ``ValueError`` with a clear message naming the source.
    Pass ``None`` for a source to skip it (e.g. in tests that only
    exercise the custom store). A missing template file or custom
    directory simply means there is nothing to collide with there.

    Comparison is case-insensitive: ControlLogix tag types are
    case-insensitive, and the custom store is one YAML file per AOI —
    on a case-insensitive filesystem (Windows) ``P_Foo.yaml`` and
    ``p_foo.yaml`` are the same file, so a name that differs only by
    case from an existing AOI must not be importable.
    """
    _validate_aoi_name(name)
    lowered = name.lower()
    if lowered in {existing.lower() for existing in _template_aoi_names(template_path)}:
        raise ValueError(
            f"AOI {name!r} already exists in the template spreadsheet "
            f"(PlantPAX library); choose a different name or re-import "
            f"it as an update to the library AOI instead"
        )
    if custom_dir is not None:
        existing = load_custom_aois(custom_dir)
        if lowered in {other.lower() for other in existing}:
            raise ValueError(
                f"AOI {name!r} already exists as a custom AOI; "
                f"re-import it to update the existing definition"
            )


def hash_file(path: str | Path) -> str:
    """SHA-256 hex digest of a file's bytes (chunked; files are ≤ 5 MB)."""
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_source_sha256(path: str | Path) -> str | None:
    """The source-file hash recorded in a custom-AOI YAML header.

    Returns ``None`` when the file is missing, unreadable, or was
    written before hashes were recorded (a legacy import) — the
    caller then falls back to a column-diff comparison instead of a
    hash comparison.
    """
    try:
        text = Path(path).read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return None
    match = _SHA256_HEADER_RE.search(text)
    return match.group(1) if match else None


def column_json(column: Column | None, subtag: str) -> dict:
    """A proposal-shaped dict for one column (JSON-safe).

    ``column`` is ``None`` when only the subtag is known — e.g. a
    removed column whose type/description lived in a proposal that is
    long gone. Type ``None`` reads as ``""``, matching the proposal
    shape.
    """
    if column is None:
        return {"subtag": subtag, "type": "", "description": "", "default": None}
    return {
        "subtag": subtag,
        "type": column.data_type or "",
        "description": column.description,
        "default": column.default,
    }


def diff_columns(before: dict[str, Column | None], after: list[Column]) -> dict:
    """Diff a confirmed custom AOI's columns against a new proposal.

    ``before`` maps subtag → the confirmed :class:`pypax.l5x.Column`,
    or ``None`` for a subtag whose type/description/default are not
    known (the source ``.L5X`` left the import folder, so no proposal
    metadata survives for it). ``after`` is the new proposal's columns
    in order. Returns ``{added, removed, changed, unchanged}``: added
    and removed are proposal-shaped column dicts, changed is
    ``[{subtag, before, after}]`` with the same dicts, unchanged is a
    count. A column is "changed" when its type, description, or
    default differs; a common subtag with unknown "before" details
    counts as unchanged — the diff reports only what it knows.
    """
    after_by_subtag = {column.subtag: column for column in after}
    added = [column_json(column, column.subtag) for column in after if column.subtag not in before]
    removed = [
        column_json(before[subtag], subtag) for subtag in before if subtag not in after_by_subtag
    ]
    changed = []
    unchanged = 0
    for subtag, new in after_by_subtag.items():
        old = before.get(subtag)
        if subtag not in before:
            continue
        if old is None:
            unchanged += 1
            continue
        if (old.data_type or "", old.description, old.default) == (
            new.data_type or "",
            new.description,
            new.default,
        ):
            unchanged += 1
        else:
            changed.append(
                {
                    "subtag": subtag,
                    "before": column_json(old, subtag),
                    "after": column_json(new, subtag),
                }
            )
    return {"added": added, "removed": removed, "changed": changed, "unchanged": unchanged}


def confirmed_meta_path(yaml_path: str | Path) -> Path:
    """The sidecar holding a confirmed AOI's column details.

    The YAML stores subtag names only; the re-import diff needs
    types/descriptions/defaults to spot "changed" columns, and the
    in-memory confirmed-metadata cache does not survive a restart (or
    a deleted source ``.L5X``). The sidecar is the durable copy,
    written next to the YAML at confirm time.
    """
    yaml_path = Path(yaml_path)
    return yaml_path.with_name(yaml_path.stem + ".meta.json")


def _write_confirmed_meta(yaml_path: Path, columns) -> None:
    """Persist confirmed column details next to the YAML.

    Best effort only: if the details cannot be serialized, no sidecar
    is written and the re-import diff falls back to added/removed (as
    before) — the YAML itself is unaffected. A column with no known
    details at all (no type, no description, no default) is not
    recorded: an entry of empty strings would read as "known and
    empty" and make the diff report spurious changes against it.
    """
    meta: dict[str, dict] = {}
    for column in columns:
        if isinstance(column, str):
            continue
        data_type = column.data_type or ""
        description = column.description or ""
        default = column.default
        if not data_type and not description and default is None:
            continue
        meta[column.subtag] = {
            "type": data_type,
            "description": description,
            "default": default,
        }
    if not meta:
        return
    try:
        text = json.dumps(meta, indent=2, sort_keys=True)
    except (TypeError, ValueError):
        return
    try:
        confirmed_meta_path(yaml_path).write_text(text + "\n", encoding="utf-8")
    except OSError:
        return


def load_confirmed_meta(yaml_path: str | Path) -> dict[str, Column]:
    """Confirmed column details from a sidecar, or ``{}`` when the
    sidecar is absent or corrupt (a legacy import, or a hand-written
    YAML). Never raises."""
    try:
        data = json.loads(confirmed_meta_path(yaml_path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if not isinstance(data, dict):
        return {}
    meta: dict[str, Column] = {}
    for subtag, entry in data.items():
        if not isinstance(subtag, str) or not isinstance(entry, dict):
            continue
        # An all-empty entry carries no details (see _write_confirmed_meta).
        if not entry.get("type") and not entry.get("description") and entry.get("default") is None:
            continue
        meta[subtag] = Column(
            subtag=subtag,
            data_type=entry.get("type") or None,
            description=entry.get("description") or "",
            default=entry.get("default"),
        )
    return meta


def save_custom_aoi(
    name: str,
    columns,
    source_l5x_name: str,
    revision: str,
    path: str | Path | None = None,
    source_sha256: str | None = None,
) -> Path:
    """Write ``config/aoi/custom/<AOI>.yaml`` for a confirmed import.

    ``columns`` is a list of :class:`pypax.l5x.Column` (or plain subtag
    strings). ``path`` may be a directory (the file is placed inside
    it) or an exact file path; the default is the repo's custom-AOI
    directory. Writing to the default location refuses a name that
    collides with a PlantPAX library AOI; overwriting the AOI's own
    file (a re-import) is allowed. ``source_sha256`` is the SHA-256 of
    the source ``.L5X`` at confirm time, recorded in the header so the
    folder scan can spot a changed export; omit it when the source
    bytes are not on disk. Returns the path written.
    """
    _validate_aoi_name(name)
    subtags = _subtags_of(columns)
    if path is None:
        check_collision(name, template_path=DEFAULT_TEMPLATE_PATH, custom_dir=None)
        target = DEFAULT_CUSTOM_DIR / f"{name}.yaml"
    else:
        target = Path(path)
        if target.suffix != ".yaml" or target.is_dir():
            target = target / f"{name}.yaml"
    header = [
        "# Custom AOI subtag definitions.",
        f"# Imported from {_header_safe(source_l5x_name)} "
        f"(TargetRevision {_header_safe(revision)}) on {date.today().isoformat()}.",
        "# Generated by pypax L5X import -- do not edit by hand;",
        "# re-import to update.",
    ]
    if source_sha256 is not None:
        header.append(f"# Source SHA-256: {source_sha256}")
    lines = [*header, f"{name}:"]
    lines.extend(f"  - '{subtag}'" for subtag in subtags)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text("\n".join(lines) + "\n", encoding="utf-8")
    _write_confirmed_meta(target, columns)
    return target


def load_custom_aois(custom_dir: str | Path | None = DEFAULT_CUSTOM_DIR) -> dict[str, list[str]]:
    """Read every custom-AOI YAML file in ``custom_dir``.

    Returns ``{AOI: [subtag, ...]}`` merged across files. A missing
    directory means no custom AOIs. Two files defining the same AOI
    raises ``ValueError``.
    """
    if custom_dir is None:
        return {}
    custom_dir = Path(custom_dir)
    if not custom_dir.is_dir():
        return {}
    definitions: dict[str, list[str]] = {}
    for yaml_path in sorted(custom_dir.glob("*.yaml")):
        for aoi, subtags in load_aoi_definitions(yaml_path).items():
            if aoi in definitions:
                raise ValueError(
                    f"custom AOI {aoi!r} is defined in both "
                    f"{yaml_path.name} and another custom file"
                )
            definitions[aoi] = subtags
    return definitions
