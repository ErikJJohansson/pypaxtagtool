# Arbiter Decisions — PyPaxTagTool webapp

Decisions Muse made as arbitrator while integrating parallel-agent
work, with the reasoning, so they can be revisited rather than
re-litigated. Erik's standing instruction: decisions are solved from
how the tool is supposed to work in its real domain, and documented
here. Earlier decisions (Phases 0–8) live in `PLAN.md` and in
`~/MEMORY.md`; this file starts at Phase 9.

## Phase 9 (Nov 24, 2026)

- **D1 — Snapshot format is JSON, not .xlsx.** The `.xlsx` export is
  the spreadsheet round-trip; a snapshot needs structure (PLC name,
  timestamp, AOI version, per-tag values) and a cheap diff, which JSON
  gives. The export keeps its own job.
- **D2 — Restore stages only, never writes.** A restore goes through
  the exact same staging path as a paste (type validation included)
  and the user still confirms the write through the normal guarded
  path (PLC-name check, typed confirmation, read-back). No second
  write path to audit.
- **D3 — A failed auto-snapshot never blocks a write.** The snapshot
  is a safety net, not a gate. Failure is logged, recorded in the
  audit entry, and reported in the write response so the UI can show
  it — but the write proceeds.
- **D4 — Pinned string sizes win over learned ones.** A learned size
  that conflicts with a pinned size is ignored and surfaced as a
  warning. Pinned sizes come from Erik's own exports; a conflicting
  learn is more likely a different UDT sharing a name.
- **D5 — Re-import is refused while staged changes exist on that
  AOI.** Updating the columns must not silently drop staged edits.
  The user discards or writes first.
- **D6 — Scan detects changed import files by SHA-256 hash**, stored
  in the custom AOI YAML metadata at confirm time. Mtime is too
  fragile across machines; the files are capped at 5 MB so hashing is
  cheap.

## Phase 9 audit (Nov 24, 2026)

Findings from the skeptical audit pass over Phase 9, and the calls
made fixing them. Each has a regression test that failed on the
pre-fix code.

- **D7 — Snapshots re-read the PLC fresh.** The auto-snapshot (and
  the manual one) captured the cached tab baseline, so a value that
  drifted since the tab loaded was snapshotted stale — and restoring
  it would silently revert the drift, the same failure mode as the
  Phase 8 audit's worst finding. Both now refresh every AOI from the
  PLC first. Cost: a full re-read before every write; correctness of
  the rollback artifact is worth it.
- **D8 — A corrupt snapshot stays downloadable and deletable.**
  Compare/restore refuse it (400), but download serves the raw bytes
  and delete removes the file without parsing it. Otherwise a corrupt
  file would be stuck in the store with no way out except by hand.
- **D9 — Confirmed column details persist in a sidecar**
  (`config/aoi/custom/<AOI>.meta.json`), not in the YAML. The YAML
  keeps its `{AOI: [subtag strings]}` shape so the existing loader
  reads it unchanged; the sidecar holds type/description/default so
  the re-import diff's "before" side survives restarts and a deleted
  source `.L5X`. A column with no known details at all is not
  recorded — an all-empty entry would read as "known and empty" and
  make the diff report spurious changes. The diff carries
  `before_details_known: false` for legacy imports that have neither
  sidecar nor cached metadata, and the UI says only added/removed
  are detectable.
- **D10 — Re-import merges learned string sizes, like a fresh
  import.** The 9b merge lived only on the fresh-import path, so a
  re-import silently dropped what its file taught. One merge path,
  warnings in both responses.
- **D11 — An upload whose filename already holds a different AOI's
  export is refused (400).** The upload is persisted before confirm,
  so overwriting it would destroy that export with no way back. Same
  AOI under the same filename is a re-import and is allowed; an
  unparseable file is not a usable export and may be replaced.
- **D12 — Snapshot names are unique across manual and auto.**
  Lookup by name is single-valued (manual wins), so a shared stem
  would leave one snapshot unreachable. Manual saves refuse a stem
  an auto-snapshot already has; auto saves suffix instead of
  shadowing a manual one; the listing dedupes the same way for
  hand-placed files.
- **D13 — Compare names the snapshot's PLC but does not refuse a
  foreign one.** The response carries `plc_name` so the UI can warn;
  blocking the compare would kill cross-PLC comparison, which the
  roadmap wants later. Restore and upload still refuse a foreign
  PLC (409) — those are the paths that could write the wrong values.
