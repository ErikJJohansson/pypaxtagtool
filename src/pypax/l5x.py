"""Parse a Studio 5000 ``.L5X`` Add-On Instruction export.

The file is treated purely as data and read with :mod:`xml.etree`
only — nothing in it is ever executed or imported. An export contains
one target AOI plus the sub-AOIs it depends on (``P_Alarm``,
``P_CmdSrc``, ...), each with ``Parameters`` and ``LocalTags``. A
member is a scalar config/status/command value; a member whose
``DataType`` names another AOI defined in the same file is a nested
sub-object (``FullStall`` → ``P_Alarm``) whose members are addressed
as dotted paths (``FullStall.Cfg_Cond``).

``curate_columns`` turns a parsed definition into the column proposal
for a new AOI tab, following Erik's rule (Nov 24, 2026, verified
against the real ``P_ValveMO_4.10.06_AOI.L5X`` export): a column is a
configuration value = its leaf name starts with ``Cfg_``. Top-level
``Cfg_{Head}AckReqd|ResetReqd|Severity`` and ``Cfg_Has{Head}Alm``
pass-throughs are dropped in favour of the nested
``{Head}.Cfg_AckReqd|Cfg_ResetReqd|Cfg_Severity|Cfg_Exists`` columns.
"""

from __future__ import annotations

import re
import struct
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

# Column order: these four come first (when present), in this order.
_FIRST_COLUMNS = ("Cfg_Desc", "Cfg_Label", "Cfg_Tag", "Cfg_Area")

# Top-level pass-throughs dropped when the nested sub-object covers
# them: Cfg_{Head}AckReqd / Cfg_{Head}ResetReqd / Cfg_{Head}Severity
# and Cfg_Has{Head}Alm map onto {Head}.Cfg_AckReqd / .Cfg_ResetReqd /
# .Cfg_Severity / .Cfg_Exists.


@dataclass(frozen=True)
class Member:
    """One AOI parameter or local tag, in L5X declaration order."""

    name: str
    data_type: str | None
    description: str = ""
    default: object = None


@dataclass
class AoiDefinition:
    """One AOI parsed from an ``.L5X`` export."""

    name: str
    revision: str = ""
    members: list[Member] = field(default_factory=list)
    sub_aois: dict[str, AoiDefinition] = field(default_factory=dict)

    def member_names(self) -> list[str]:
        return [member.name for member in self.members]


@dataclass(frozen=True)
class Column:
    """One proposed tab column: a config value of an AOI."""

    subtag: str  # dotted, with leading dot: ".Cfg_Desc", ".FullStall.Cfg_Cond"
    data_type: str | None
    description: str = ""
    default: object = None


def _decode_default(element: ET.Element, data_type: str | None) -> object:
    """Decode a member's ``DefaultData`` per its data type.

    Returns ``None`` when the member has no default (or the default
    cannot be decoded — a bad default must not fail the whole import).
    BOOL is one byte (0/1); DINT and REAL are 4-byte little-endian;
    STRING_* is a 4-byte little-endian length followed by that many
    ASCII bytes. ``Format="String"`` defaults are quoted
    (``'Valve Control'``). Hex text may contain whitespace/newlines.
    """
    default_el = element.find("DefaultData")
    if default_el is None:
        return None
    if default_el.get("Format") == "String":
        text = "".join(default_el.itertext()).strip()
        match = re.search(r"'([^']*)'", text)
        return match.group(1) if match else (text or None)
    # Decorated defaults carry DataValue children instead of text.
    if default_el.find("DataValue") is not None:
        value_el = default_el.find("DataValue")
        if value_el is not None and value_el.get("Value") is not None:
            return value_el.get("Value")
    raw = "".join(default_el.itertext()).strip()
    if not raw:
        return None
    try:
        data = bytes.fromhex(re.sub(r"\s+", "", raw))
    except ValueError:
        return None
    if not data:
        return None
    try:
        if data_type == "BOOL":
            return bool(data[0])
        if data_type == "DINT":
            return struct.unpack("<i", data[:4])[0]
        if data_type == "REAL":
            return struct.unpack("<f", data[:4])[0]
        if data_type and data_type.startswith("STRING"):
            length = struct.unpack("<I", data[:4])[0] if len(data) >= 4 else 0
            return data[4 : 4 + length].decode("ascii", errors="ignore")
    except (struct.error, UnicodeDecodeError):
        return None
    return None


def _parse_members(aoi_el: ET.Element) -> list[Member]:
    """Parameters then local tags, in declaration order."""
    members: list[Member] = []
    seen: set[str] = set()
    for section in ("Parameters", "LocalTags"):
        section_el = aoi_el.find(section)
        if section_el is None:
            continue
        for child in section_el:
            name = child.get("Name")
            if not name or name in seen:
                continue
            seen.add(name)
            description_el = child.find("Description")
            description = (
                "".join(description_el.itertext()).strip() if description_el is not None else ""
            )
            members.append(
                Member(
                    name=name,
                    data_type=child.get("DataType"),
                    description=description,
                    default=_decode_default(child, child.get("DataType")),
                )
            )
    return members


def _revision_of(root: ET.Element, aoi_el: ET.Element) -> str:
    revision = aoi_el.get("Revision") or ""
    extension = aoi_el.get("RevisionExtension") or ""
    if revision:
        return f"{revision}{extension}".strip()
    return (root.get("TargetRevision") or "").strip()


def parse_l5x(path: str | Path) -> AoiDefinition:
    """Parse a Studio 5000 ``.L5X`` AOI export (read-only, XML only).

    Returns the target AOI (the file's ``TargetName``) with every other
    AOI defined in the same file attached as ``sub_aois``. Raises
    ``ValueError`` for a file that is not a readable AOI export.
    """
    path = Path(path)
    try:
        tree = ET.parse(path)
    except ET.ParseError as exc:
        raise ValueError(f"{path}: not a readable .L5X file: {exc}") from exc
    except OSError as exc:
        raise ValueError(f"{path}: cannot read .L5X file: {exc}") from exc
    root = tree.getroot()
    definitions: dict[str, AoiDefinition] = {}
    order: list[str] = []
    for aoi_el in root.iter("AddOnInstructionDefinition"):
        name = aoi_el.get("Name")
        if not name or name in definitions:
            continue
        definitions[name] = AoiDefinition(
            name=name,
            revision=_revision_of(root, aoi_el),
            members=_parse_members(aoi_el),
        )
        order.append(name)
    if not definitions:
        raise ValueError(f"{path}: no AddOnInstructionDefinition found")
    target_name = root.get("TargetName") or order[0]
    target = definitions.get(target_name)
    if target is None:
        raise ValueError(f"{path}: TargetName {target_name!r} has no definition in the file")
    target.sub_aois = {name: definitions[name] for name in order if name != target.name}
    return target


def learn_string_sizes(root: ET.Element) -> dict[str, int]:
    """Learn string-UDT sizes from an ``.L5X`` export's type table.

    Scans ``<DataType Family="StringFamily">`` blocks for the ``DATA``
    member: a string UDT holds ``Dimension`` SINT characters, so that
    number is the type's character limit. Keys are upper-cased type
    names (``STRING_COND``). Only StringFamily UDTs with a SINT
    ``DATA`` member count — a non-string type, a StringFamily UDT
    with no ``DATA`` member, or a ``DATA`` member that is not SINT /
    has no usable dimension is ignored, never guessed. When the same
    type is defined twice, the last definition wins.
    """
    sizes: dict[str, int] = {}
    for data_type in root.iter("DataType"):
        if data_type.get("Family") != "StringFamily":
            continue
        name = (data_type.get("Name") or "").strip()
        if not name:
            continue
        members = data_type.find("Members")
        if members is None:
            continue
        for member in members:
            if member.get("Name") != "DATA" or member.get("DataType") != "SINT":
                continue
            try:
                size = int(member.get("Dimension") or "")
            except ValueError:
                continue
            if size > 0:
                sizes[name.upper()] = size
    return sizes


def _nested_heads(definition: AoiDefinition) -> dict[str, AoiDefinition]:
    """Members whose DataType is another AOI defined in the same file."""
    heads: dict[str, AoiDefinition] = {}
    for member in definition.members:
        if member.data_type in definition.sub_aois:
            heads[member.name] = definition.sub_aois[member.data_type]
    return heads


def _passthrough_names(heads: dict[str, AoiDefinition]) -> set[str]:
    """Top-level pass-throughs covered by a nested sub-object.

    A pass-through is only dropped when the nested member it maps to
    actually exists, so a config value is never lost silently.
    """
    dropped: set[str] = set()
    for head, sub in heads.items():
        sub_names = {member.name for member in sub.members}
        candidates = {
            f"Cfg_{head}AckReqd": "Cfg_AckReqd",
            f"Cfg_{head}ResetReqd": "Cfg_ResetReqd",
            f"Cfg_{head}Severity": "Cfg_Severity",
            f"Cfg_Has{head}Alm": "Cfg_Exists",
        }
        for top_name, nested_name in candidates.items():
            if nested_name in sub_names:
                dropped.add(top_name)
    return dropped


def curate_columns(definition: AoiDefinition) -> list[Column]:
    """Propose the config columns for ``definition``'s tab.

    A column is a config value = its leaf name starts with ``Cfg_``.
    Top-level pass-throughs covered by a nested sub-object are
    dropped; nested sub-objects expand into dotted columns. Order:
    ``Cfg_Desc`` / ``Cfg_Label`` / ``Cfg_Tag`` / ``Cfg_Area`` first
    (when present), then the rest in L5X declaration order, with a
    nested sub-object's columns at the position of its head member.
    """
    heads = _nested_heads(definition)
    dropped = _passthrough_names(heads)

    ordered: list[Column] = []

    def _columns_for(member: Member) -> list[Column]:
        if member.name in heads:
            sub = heads[member.name]
            return [
                Column(
                    subtag=f".{member.name}.{leaf.name}",
                    data_type=leaf.data_type,
                    description=leaf.description,
                    default=leaf.default,
                )
                for leaf in sub.members
                if leaf.name.startswith("Cfg_")
            ]
        if member.name.startswith("Cfg_") and member.name not in dropped:
            return [
                Column(
                    subtag=f".{member.name}",
                    data_type=member.data_type,
                    description=member.description,
                    default=member.default,
                )
            ]
        return []

    for member in definition.members:
        ordered.extend(_columns_for(member))

    first = [col for col in ordered if col.subtag[1:] in _FIRST_COLUMNS]
    first.sort(key=lambda col: _FIRST_COLUMNS.index(col.subtag[1:]))
    rest = [col for col in ordered if col.subtag[1:] not in _FIRST_COLUMNS]
    return first + rest
