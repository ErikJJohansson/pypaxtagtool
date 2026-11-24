"""Audit log: in-memory entries plus one JSON-lines file per day.

Every entry is a dict with ``timestamp`` (UTC ISO), ``plc_name``,
``action``, ``tags``, and ``results``. Entries are always kept in
memory (what ``GET /api/audit`` returns); when an audit directory is
configured they are also appended to ``audit-YYYY-MM-DD.jsonl`` in
that directory, one JSON object per line, so a write is still on disk
even if the app is restarted. File writes are best-effort: a failed
append never breaks the request that produced the entry.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path


class AuditLog:
    """In-memory audit entries, mirrored to daily JSONL files."""

    def __init__(self, audit_dir=None):
        # None means "audit/" under the current working directory (the
        # repo root when launched with pypax-web from there). Tests pass
        # a tmp_path so the suite never writes into the repo.
        self.audit_dir = Path(audit_dir) if audit_dir is not None else Path.cwd() / "audit"
        self.entries: list[dict] = []

    def record(self, entry):
        """Append an entry to memory and to today's JSONL file."""
        entry = dict(entry)
        self.entries.append(entry)
        try:
            self.audit_dir.mkdir(parents=True, exist_ok=True)
            day = datetime.now(timezone.utc).strftime("%Y-%m-%d")
            path = self.audit_dir / f"audit-{day}.jsonl"
            with path.open("a", encoding="utf-8") as fh:
                fh.write(json.dumps(entry, sort_keys=True) + "\n")
        except OSError:
            pass  # best-effort: the in-memory entry is the record

    def list(self):
        """A copy of the in-memory entries, oldest first."""
        return list(self.entries)
