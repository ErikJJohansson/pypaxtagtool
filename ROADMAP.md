# PyPaxTagTool — Feature Roadmap

Where the webapp stands (Nov 24, 2026): Phases 0–8 done — core
refactor, PLC optimization, FastAPI backend, AG Grid frontend, write
safety with read-back verification, spreadsheet export, docs, custom
AOI import via L5X — plus an adversarial audit pass (11 findings
fixed) and per-type data validation with string-length checks. 235
tests, ruff clean, all on the local `feature/webapp` branch; `main`
untouched; no live PLC tested yet (see Milestones).

This is a proposed order, not a commitment. Erik reorders it — move
things up, cut things, add things.

## Next up

### 1. PLC snapshots — save, compare, restore (asked Nov 24, 2026)
**DONE Nov 24, 2026 (Phase 9a + audit).** Named JSON snapshots with
PLC name/timestamp/AOI version, compare vs live PLC, restore by
staging through the guarded write path, upload from another machine,
auto-snapshot before every write (fresh pre-write values, newest 10
per PLC).
Today the only snapshot is the `.xlsx` export: it holds baseline
values and the CLI can write one back, but it has no name, no
timestamp, no PLC metadata, no compare, no restore in the UI, and it
skips custom AOIs. A real snapshot feature:
- **Save:** capture the live baseline for every AOI (template +
  custom) into a named JSON snapshot — PLC name, timestamp, AOI
  version, all values. Stored under `snapshots/` (gitignored);
  downloadable as a file.
- **Compare:** diff a snapshot against the live PLC (or against
  another snapshot) and show what changed, per tag: snapshot value
  vs current value.
- **Restore:** stage the snapshot's values through the normal
  guarded write path — PLC-name check, typed confirmation, chunked
  writes, per-tag read-back. Restore is a staged write, not a
  separate write path. A snapshot file can also be uploaded from
  another machine and restored the same way.
- **Auto-snapshot before every write:** capture the pre-write state
  automatically, so any write can be rolled back from its own
  snapshot. This is the safety half of the feature.

### 2. Learn string-UDT sizes from L5X imports
**DONE Nov 24, 2026 (Phase 9b + audit).** The parser reads
`<DataType>` blocks and learns string sizes into
`config/string_udt_sizes.yaml`; pinned sizes win on conflict.
The validator pins string limits from exports (STRING_Area 8 from
P_ValveMO, STRING_Cond / STRING_NavTag 20 from P_Intlk). Teach the
L5X parser to read the `<DataType>` block of any imported file and
learn the DATA dimension of every string UDT it defines, so a future
custom string type gets a limit automatically instead of waiting for
a hardcoded pin.

### 3. Re-import a custom AOI in place
**DONE Nov 24, 2026 (Phase 9c + audit).** A re-upload shows a diff
and updates the YAML in place; refused while staged changes exist
on that AOI; scan reports changed files (SHA-256).
Today updating a custom AOI means delete + re-import, and the folder
scan skips already-imported AOIs. Instead: a re-import shows a diff
(added / removed / renamed columns) and confirms to update the AOI
in place; the scan proposes updates for AOIs whose file changed.

### 4. Profile picker in the connect form
The connect form has a free-text profile field. Replace it with a
dropdown of the configured profiles (from `config.toml` /
`config/plcs/`), showing each profile's safety level, so picking the
wrong PLC's profile gets harder.

## Decided against

- **Exporting custom AOIs** (cut Nov 24, 2026, Erik's call): a custom
  AOI requires its logic — the AOI definition lives in Studio 5000 —
  and the `.xlsx` export would only carry its data. The spreadsheet
  round-trip is the PlantPAX template's config tool; it doesn't apply
  to a custom AOI, so custom AOIs stay out of the export
  deliberately, not as a missing feature.

## Later

### 6. Packaged app (PyInstaller, macOS + Windows)
One-file builds that run on a plant laptop with no Python install —
the tool is meant to be carried to the PLC, not just run from a dev
machine. Same pure-Python stack, just bundled.

### 7. Tag / instance search across AOIs
Find an instance or subtag by name from anywhere in the app, jump
to its tab and row. Today you have to know which tab it's on.

### 8. Bulk tools: copy instance, fill down, find-and-replace
Copy one instance's config to several others, fill a column down a
range, replace a value across an AOI — with the same type validation
and staging as a paste.

### 9. Per-tag change history
From snapshots + the audit log: for any tag, what changed, when, and
to what. (Snapshots in §1 are the raw material; this is the view.)

### 10. Custom-AOI-only mode
Today a template workbook is still required even if you only use
custom AOIs. Allow running with custom AOIs alone.

### 11. Read-only mode
A launch flag that disables writes entirely — safe lookaround on a
live plant, and a sane default for a first look at an unfamiliar PLC.

### 12. Compare two PLCs (or two snapshots) side by side
Snapshot §1 covers one PLC vs its past. This extends it: diff the
same AOI across two PLCs, or two snapshot files, in one view.

### 13. Whole-app backup / restore
One file (or folder) holding config, custom AOIs, and snapshots, so
the setup moves between machines without re-importing everything.

## Milestones (not features)

- **First live PLC smoke test** (`docs/SMOKE_TEST.md`) — read-only
  first, then one harmless single-tag write. The only thing no
  fake-PLC test can close, and it gates trusting the write path.
- **Merge `feature/webapp` → `main`** (Erik's call) and a v1.0
  release note. Includes the P_ValveMO typo fix and Phase 8, which
  are still local-only by Erik's standing instruction.
- **GitHub Actions CI** — blocked on the Muse GitHub App lacking
  the Workflows permission; the workflow file exists locally only
  and must be added manually on GitHub.
