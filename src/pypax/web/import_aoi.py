"""L5X import for custom AOI templates (Phase 8b).

An import starts from a Studio 5000 ``.L5X`` export — either uploaded
in the UI or found by scanning the import directory — and produces an
:class:`ImportProposal`: the AOI name, revision, curated config columns
(see :func:`pypax.l5x.curate_columns`), and the nested sub-objects
found. A proposal is inert data: nothing is writable from it. Only a
confirmed proposal becomes a tab, via
:func:`pypax.custom_aoi.save_custom_aoi`.

The uploaded bytes are treated purely as data: they are parsed with
:mod:`xml.etree` only (via :func:`pypax.l5x.parse_l5x`), never executed
or imported. Multipart bodies are parsed with the stdlib ``email``
parser — the project has no ``python-multipart`` dependency and the
contract forbids adding one.

Proposals live server-side in memory, keyed by ``proposal_id``. A
proposal is consumed on successful confirm: confirming the same id
again is an unknown proposal, not a collision. A proposal for an
already-imported AOI name is a re-import: the upload/scan answers
with ``status: "update_available"`` and a column diff, and confirming
it rewrites the custom AOI in place (all-or-nothing, no column
checkboxes). Column metadata (types/descriptions) is kept separately in
``column_meta`` so confirmed AOIs still serve typed columns after
their proposal is consumed.
"""

from __future__ import annotations

import os
import re
import tempfile
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from email.parser import BytesParser
from email.policy import default as _email_default_policy
from pathlib import Path

from ..l5x import AoiDefinition, Column, curate_columns, learn_string_sizes, parse_l5x

MAX_UPLOAD_BYTES = 5 * 1024 * 1024  # 5 MB cap on the uploaded file

REPO_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_IMPORT_DIR_NAME = "aoi_imports"

_AOI_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def valid_aoi_name(name: str) -> bool:
    """True when ``name`` is a usable AOI identifier."""
    return isinstance(name, str) and bool(_AOI_NAME_RE.match(name))


def resolve_import_dir(settings, explicit: str | Path | None = None) -> Path:
    """Import directory: explicit path, else the setting, else ``aoi_imports/``.

    Relative paths resolve against the repo root (the task's default);
    absolute paths are used as-is.
    """
    if explicit is not None:
        return Path(explicit)
    raw = (settings.aoi_import_dir if settings is not None else None) or DEFAULT_IMPORT_DIR_NAME
    path = Path(raw)
    return path if path.is_absolute() else REPO_ROOT / path


@dataclass
class ImportProposal:
    """One parsed, unconfirmed AOI import."""

    proposal_id: str
    aoi: str
    revision: str
    source_file: str  # basename, for display
    source_path: str | None  # full path when the file lives on disk
    columns: list[Column] = field(default_factory=list)
    nested_objects: list[str] = field(default_factory=list)
    # String-UDT sizes learned from the file's type table
    # ({NAME: size}); merged into the registry at confirm time.
    learned_sizes: dict[str, int] = field(default_factory=dict)

    def to_json(self) -> dict:
        """The frozen proposal shape the frontend builds against."""
        return {
            "aoi": self.aoi,
            "revision": self.revision,
            "source_file": self.source_file,
            "columns": [
                {
                    "subtag": column.subtag,
                    "type": column.data_type or "",
                    "description": column.description,
                    "default": column.default,
                    "selected": True,
                }
                for column in self.columns
            ],
            "nested_objects": list(self.nested_objects),
            "proposal_id": self.proposal_id,
        }


def nested_object_names(definition: AoiDefinition) -> list[str]:
    """Head members whose DataType is another AOI defined in the file."""
    return [member.name for member in definition.members if member.data_type in definition.sub_aois]


def build_proposal(
    definition: AoiDefinition,
    source_file: str,
    source_path: str | Path | None = None,
    proposal_id: str | None = None,
) -> ImportProposal:
    """Curate ``definition`` into a proposal (fresh id unless given).

    String-UDT sizes are learned from the source file's type table
    when the file is on disk; a file that cannot be re-read yields no
    learned sizes (the proposal itself is unaffected).
    """
    learned_sizes: dict[str, int] = {}
    if source_path is not None:
        try:
            learned_sizes = learn_string_sizes(ET.parse(source_path).getroot())
        except (OSError, ET.ParseError):
            learned_sizes = {}
    return ImportProposal(
        proposal_id=proposal_id or uuid.uuid4().hex,
        aoi=definition.name,
        revision=definition.revision,
        source_file=source_file,
        source_path=str(source_path) if source_path is not None else None,
        columns=curate_columns(definition),
        nested_objects=nested_object_names(definition),
        learned_sizes=learned_sizes,
    )


def parse_l5x_bytes(data: bytes, filename: str) -> AoiDefinition:
    """Parse uploaded ``.L5X`` bytes (never executed, XML only).

    ``parse_l5x`` reads from a path, so the bytes go through a temp
    file that is removed immediately after parsing. Raises
    ``ValueError`` for anything that is not a readable AOI export.
    """
    if len(data) > MAX_UPLOAD_BYTES:
        raise ValueError(f"upload exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB cap")
    fd, tmp_path = tempfile.mkstemp(suffix=".L5X")
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(data)
        try:
            return parse_l5x(tmp_path)
        except ValueError as exc:
            raise ValueError(f"{filename}: {exc}") from exc
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def parse_multipart_upload(body: bytes, content_type: str) -> tuple[str, bytes]:
    """Extract the ``file`` part from a multipart/form-data body.

    Stdlib ``email`` parser only — no ``python-multipart``. Returns
    ``(filename, file_bytes)``. Raises ``ValueError`` for a body that
    is not multipart, has no ``file`` part, or has no filename.
    """
    if not body:
        raise ValueError("empty upload")
    if len(body) > MAX_UPLOAD_BYTES + 65536:  # body carries multipart overhead too
        raise ValueError(f"upload exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB cap")
    if "\r" in content_type or "\n" in content_type:
        raise ValueError("invalid Content-Type header")
    header = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode(
        "ascii", errors="ignore"
    )
    message = BytesParser(policy=_email_default_policy).parsebytes(header + body)
    if not message.is_multipart():
        raise ValueError("expected a multipart/form-data upload")
    for part in message.iter_parts():
        if part.get_param("name", header="Content-Disposition") != "file":
            continue
        filename = part.get_filename() or ""
        payload = part.get_payload(decode=True)
        return filename, payload if payload is not None else b""
    raise ValueError("no file part in upload")


def check_upload_target(import_dir: str | Path, filename: str, aoi_name: str) -> None:
    """Refuse an upload whose filename already holds a different AOI.

    The upload is persisted into the import directory before confirm,
    so overwriting a file that holds another AOI's export would
    destroy it with no way back. A file holding the same AOI is a
    re-import and is allowed; an unparseable file (or one over the
    import cap) is not a usable export and may be replaced. Raises
    ``ValueError``.
    """
    base = Path(filename or "").name
    if not base:
        return
    target = Path(import_dir) / base
    if not target.is_file():
        return
    try:
        if target.stat().st_size > MAX_UPLOAD_BYTES:
            return
        existing = parse_l5x(target)
    except (ValueError, OSError):
        return
    if existing.name != aoi_name:
        raise ValueError(
            f"upload {base!r} holds AOI {aoi_name!r} but the file already "
            f"holds a different AOI {existing.name!r}; rename the upload "
            f"or delete the existing file first"
        )


def save_upload(import_dir: str | Path, filename: str, data: bytes) -> Path:
    """Persist an uploaded ``.L5X`` into the import directory.

    The basename is used (any directory in the uploaded name is
    dropped), so an upload can never write outside the import dir.
    Raises ``ValueError`` for an empty or unsafe filename.
    """
    base = Path(filename or "").name
    if not base or base in (".", ".."):
        raise ValueError("upload has no usable filename")
    if any(ord(ch) < 0x20 or ord(ch) == 0x7F for ch in base):
        raise ValueError("upload filename contains control characters")
    target = Path(import_dir) / base
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(data)
    return target


@dataclass
class ScanResult:
    """One file found by a folder scan: parsed definition or error."""

    path: Path
    definition: AoiDefinition | None = None
    error: str | None = None


def iter_l5x_files(import_dir: str | Path) -> list[Path]:
    """``.L5X`` files in ``import_dir`` (case-insensitive), sorted by name."""
    directory = Path(import_dir)
    if not directory.is_dir():
        return []
    return sorted(
        (
            entry
            for entry in directory.iterdir()
            if entry.is_file() and entry.suffix.lower() == ".l5x"
        ),
        key=lambda entry: entry.name.lower(),
    )


def scan_import_dir(import_dir: str | Path) -> list[ScanResult]:
    """Parse every ``.L5X`` in the import directory.

    A file that cannot be parsed yields a ``ScanResult`` with ``error``
    set instead of failing the whole scan. A file over the import cap
    is skipped the same way (the upload path enforces the same cap, so
    the scan must not let an unbounded file in through the folder).
    Never raises for a missing directory (there is simply nothing to
    import yet).
    """
    results: list[ScanResult] = []
    for path in iter_l5x_files(import_dir):
        try:
            if path.stat().st_size > MAX_UPLOAD_BYTES:
                raise ValueError(
                    f"file exceeds the {MAX_UPLOAD_BYTES // (1024 * 1024)} MB import cap"
                )
            results.append(ScanResult(path=path, definition=parse_l5x(path)))
        except ValueError as exc:
            results.append(ScanResult(path=path, error=str(exc)))
    return results


def register_proposal(state, proposal: ImportProposal) -> ImportProposal:
    """Store ``proposal`` in the app state, reusing the id for its path.

    A rescan (or re-upload) of the same file refreshes the stored
    proposal in place and keeps its ``proposal_id`` stable, so a UI
    holding the id is not invalidated by a rescan. Column metadata is
    cached separately and survives confirm (proposals are consumed).
    Returns the stored proposal.
    """
    if proposal.source_path is not None:
        existing_id = state.proposals_by_path.get(proposal.source_path)
        if existing_id is not None and existing_id in state.proposals:
            proposal.proposal_id = existing_id
    state.proposals[proposal.proposal_id] = proposal
    if proposal.source_path is not None:
        state.proposals_by_path[proposal.source_path] = proposal.proposal_id
    state.column_meta[proposal.aoi] = {column.subtag: column for column in proposal.columns}
    return proposal


def consume_proposal(state, proposal_id: str) -> ImportProposal | None:
    """Remove a proposal from the confirmable store (on successful confirm).

    Returns the removed proposal, or ``None`` when the id is unknown.
    Column metadata is deliberately kept: confirmed AOIs still serve
    typed columns after their proposal is gone.
    """
    proposal = state.proposals.pop(proposal_id, None)
    if proposal is not None and proposal.source_path is not None:
        state.proposals_by_path.pop(proposal.source_path, None)
    return proposal
