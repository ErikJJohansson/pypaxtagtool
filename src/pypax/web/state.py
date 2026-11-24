"""Shared mutable state for one webapp instance.

All request handlers read and write through a single :class:`AppState`
guarded by one :class:`threading.Lock`: the app talks to one PLC at a
time, and a write must never race a read. The audit log is kept in
memory and mirrored to daily JSONL files (see :mod:`pypax.web.audit`);
everything else is in-memory only.
"""

from __future__ import annotations

import threading
from typing import Any

from ..config import AppSettings
from .audit import AuditLog


class AppState:
    """Connection, baseline values, staged changes, and the audit log."""

    def __init__(self, plc_factory, template_path, settings, profiles, audit_dir=None):
        self.lock = threading.Lock()
        self.plc_factory = plc_factory
        self.template_path = template_path  # resolved path or None
        self.settings: AppSettings = settings
        self.profiles = profiles  # dict[str, PlcProfile]

        self.plc = None
        self.commpath: str | None = None
        self.plc_name: str | None = None
        self.profile = None  # PlcProfile | None

        # Baseline: last-read PLC values, aoi -> instance -> {subtag: value}.
        self.baseline: dict[str, dict[str, dict[str, Any]]] = {}
        # Staged changes: (aoi, instance, subtag) -> change dict with
        # aoi/instance/subtag/tag/old_value/new_value.
        self.staged: dict[tuple[str, str, str], dict[str, Any]] = {}
        # Audit log: one entry per connect/disconnect/write, kept in
        # memory and mirrored to daily JSONL files.
        self.audit = AuditLog(audit_dir)
        # Cached AOI model from the template workbook (or None).
        self.aoi_model = None
        # Custom AOIs (Phase 8): names present in the model that came
        # from config/aoi/custom/*.yaml rather than the template.
        self.custom_aoi_names: set[str] = set()
        # Custom-AOI storage and L5X import state (set by create_app).
        self.custom_aoi_dir = None  # Path | None; None means the repo default
        self.import_dir = None  # Path | None
        # String-UDT size registry file (set by create_app); None means
        # the repo default (config/string_udt_sizes.yaml).
        self.string_registry_path = None  # Path | None
        # Snapshot storage (Phase 9a), set by create_app.
        self.snapshot_dir = None  # Path | None
        # Confirmable import proposals, keyed by proposal_id; a proposal
        # is consumed (removed) on successful confirm.
        self.proposals: dict[str, object] = {}
        # Imported file path -> proposal_id (rescan keeps the id stable).
        self.proposals_by_path: dict[str, str] = {}
        # Column metadata (type/description) per AOI, cached from
        # proposals. Kept after confirm so confirmed AOIs still serve
        # typed columns; dropped when the custom AOI is deleted.
        self.column_meta: dict[str, dict[str, object]] = {}
        # Confirmed column metadata per AOI: written only at confirm
        # time (fresh import or re-import), never by a proposal
        # registration — so a re-import diff always compares against
        # the confirmed definition, not the latest unconfirmed
        # proposal. Seeded at startup when the file on disk still
        # matches the hash stored in the YAML. Used only for diffing.
        self.confirmed_meta: dict[str, dict[str, object]] = {}

    @property
    def connected(self) -> bool:
        return self.plc is not None

    def reset_connection(self) -> None:
        """Drop the connection and everything derived from it (not the audit)."""
        self.plc = None
        self.commpath = None
        self.plc_name = None
        self.profile = None
        self.baseline = {}
        self.staged = {}
        self.aoi_model = None

    def drop_aoi_data(self, aoi: str) -> None:
        """Drop baseline + staged changes for one AOI (re-import/delete).

        Column metadata is kept: a re-imported AOI re-enriches from its
        new proposal, and a deleted AOI drops its metadata separately.
        """
        self.baseline.pop(aoi, None)
        for key in [key for key in self.staged if key[0] == aoi]:
            del self.staged[key]
