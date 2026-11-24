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
