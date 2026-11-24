# PyPaxTagTool Webapp — Plan (v2, final for approval)

Repo: `ErikJJohansson/pypaxtagtool` · Branch: `feature/webapp` off `main` (after PR #6 merged as `b76db34`)
Target: PlantPAX v4.10.6 · Template: `ProcessLibraryOnlineConfigTool.xlsm` only

## 1. Goal

Turn the CLI-only `pypax.py` into a webapp that runs locally on Erik's computer,
connects to a PLC over EtherNet/IP via pycomm3, and lets a user read PlantPAX AOI
config tags, edit them in a spreadsheet-like grid, and write changes back —
with the spreadsheet workflow still working exactly as today.

## 2. Scope / non-scope

In scope:
- Webapp (FastAPI backend + browser grid), same core library as the CLI
- One tab per AOI, copy-paste from spreadsheets, changed-cell highlighting
- Diff-only, chunked, read-back-verified PLC writes; PLC-name double-check
- Python + PLC speed optimizations, backwards-compatible spreadsheet round-trip
- Unit + integration tests (fake PLC), CI

Out of scope:
- `SpotConfigTool.xlsm` — ignore it entirely
- No site-specific config in the repo (it is public): no PLC IPs, PLC names, plant/site names
- No live-PLC testing by Muse — no PLC is available; pycomm3 itself is proven/working
- No multi-user / auth / remote hosting in v1 — localhost only
- Cross-platform is a hard requirement: Erik used the CLI on his Mac; escaping
  Excel VBA + RSLinx Classic (Windows-only) is the reason the tool is Python.
  Pure-Python stack only (pycomm3, openpyxl, FastAPI/uvicorn — all cross-platform),
  no COM / RSLinx / win32-only deps, `pathlib` everywhere, no hardcoded drive
  letters or backslashes. Packaging is `pip install` + a launch command that works
  identically on macOS and Windows; CI runs the test matrix on macOS + Windows (+ Linux).

## 3. Current state (verified 2026-11-24)

- `main` at `b76db34` (PR #6 squash-merged). `pypax.py` 450 lines, CLI only:
  `pypax.py <ip/slot> read|write [file]`. Deps: pycomm3, openpyxl, tqdm — all unpinned.
- `AOI_definitions.py`: 57 AOIs, header says v4.10.4. Template has 122 sheets,
  56 AOI sheets; definitions match the spreadsheet on sampled AOIs
  (P_AIn 101/101, P_Motor 81/81, P_PIDE 185/185).
- Layout: data types row 6, subtag rows 7–8, defaults row 9, instances from row 10,
  tag name col C, values from col E; Setup sheet holds instance counts.
- No tests, no CI. `.gitignore` only covers `*.xlsx` / `*.xls` / `.DS_Store`.
- Known weaknesses: one `plc.read()` per instance row; PLC tag list rescanned once
  per AOI; set-based diff (loses order, collapses duplicates, brittle float equality);
  one giant unchunked `plc.write()`; no read-back confirmation; per-cell openpyxl access;
  bare `except:`s; `AOI_definitions.py` is a 96 KB Python dict.

## 4. Architecture

- Backend: FastAPI, runs locally, localhost only. One persistent `LogixDriver`
  connection behind a lock; no parallel writes to one PLC.
- Frontend: browser grid on AG Grid Community (MIT, native Excel TSV paste),
  one tab per AOI (56), lazy-loaded, instance counts from the Setup sheet.
- Core: one shared library used by both CLI and webapp, so the two cannot drift apart.
- Stack: Python + FastAPI + pycomm3 + openpyxl, pytest + ruff, GitHub Actions CI.

## 5. Config organization

Three layers, one source of truth each; real values never committed.

- Spreadsheet layout (rows/cols, Setup columns): one layout config, read by CLI,
  webapp, and tests — no duplicated magic constants.
- AOI definitions: data files, versioned — `config/aoi/v4.10.6.yaml` (and `v4.10.4`
  kept for legacy). Generated from the spreadsheet by `scripts/extract_aoi_definitions.py`;
  a test asserts the checked-in file matches the spreadsheet. `AOI_definitions.py`
  stays as a thin compatibility shim.
- App / connection policy: `config.example.toml` in the repo; real `config.toml`
  gitignored on the machine. Pydantic Settings, `PYPAX_` env overrides, fail fast
  at startup on missing/invalid config or AOI-version mismatch. Holds: comm path,
  expected PLC name, timeouts, write chunk size, REAL read-back tolerance, retries.
- Per-PLC profiles (`config/plcs/<name>.toml`, local only): comm path, expected name,
  AOI version, safety level. `test` profiles write freely; `production` profiles
  require the name double-check + typed confirmation. The app refuses to write when
  the profile's expected name ≠ the PLC it actually connected to.
- Dependencies: `pyproject.toml` with pinned versions + lock file; dev deps
  (pytest, ruff) separated. `.gitignore` extended: `__pycache__/`, `.venv/`,
  `config.toml`, `config/plcs/`, `*.log`, audit logs. Template `.xlsm` moves to
  Git LFS or a release asset; test fixtures are small generated workbooks.

## 6. Phases (small commits, one chunk each)

Status (2026-11-24): Phases 0–7 done and merged into `feature/webapp`
(177 tests passing). `main` untouched at `b76db34` — nothing merges to
main without Erik's explicit say-so. Phase 8 specced Nov 24, 2026 (below),
implementation in progress.

- Phase 0 — Scaffolding. Branch `feature/webapp`; pytest + GitHub Actions CI
  (test matrix: macOS + Windows + Linux). NOTE (2026-11-24): the Muse GitHub App
  has no Workflows permission, so `.github/workflows/ci.yml` cannot be pushed via
  the API — it lives in the local branch only, ignored for now; local pytest + ruff
  are the gate until it is added manually on GitHub later.
  repo layout (`src/pypax/`); fake-PLC harness (in-memory `LogixDriver` stand-in
  with the same interface: tags, read, write, `get_plc_name`).
- Phase 1 — Core refactor + unit tests first. Extract pure logic from `pypax.py`
  (spreadsheet parsing, AOI/subtag handling, diffing, type coercion) into
  `src/pypax/`; `pypax.py` becomes a thin CLI shim with identical behaviour.
  Golden test: CLI read/write on a small generated workbook is byte-equivalent
  before/after. Light v4.10.6 audit here: definitions-vs-spreadsheet test +
  spot-checks (per Erik, 4.10.4 → 4.10.6 is logic-only, no tag changes).
- Phase 2 — PLC + Python speed. Index PLC tags once by data type; bulk-read per
  AOI in CIP-sized chunks instead of one read per row; cache subtag lists; row
  iteration instead of per-cell openpyxl access; remove `tag_list + ...` and
  per-AOI rescans; replace set-diff with ordered, deduped, tolerance-aware diff.
- Phase 3 — Webapp backend. FastAPI: connect/disconnect, read-all, per-AOI data,
  staged changes, write, read-back verification, audit log. Same core as the CLI.
- Phase 4 — Frontend. Tabs per AOI; AG Grid; TSV paste mapped to
  (instance, subtag) cells and type-checked against row 6 (bad values flagged red,
  never silently accepted); edited cells highlight amber vs the last-read PLC baseline.
- Phase 5 — Write safety + confirmation. Banner shows `get_plc_name()` on connect;
  pre-write re-check of name + profile; typed confirmation listing exact tag counts
  per AOI; chunked sequential writes; read-back of every written tag with per-tag
  verdict (verified / mismatch / failed); retry failed once; audit log per write.
- Phase 6 — Spreadsheet round-trip. Export to `.xlsx` in the exact template layout
  (headers rows 4–9, instances row 10+, Setup counts) so webapp files work with the
  CLI and vice versa; round-trip test.
- Phase 7 — Docs + handover. README (run it, point it at a PLC, what to do when a
  write fails verification), live smoke-test checklist (read-only first, compare
  against known PLC values, then one harmless single-tag write), release notes.
- Phase 8 — Custom AOI templates (L5X import + folder scan). Specced Nov 24,
  2026; DONE Nov 24, 2026 — 8a parser+store (feccf16), 8b backend (ea2e9a2),
  8c frontend (ca4f23c), 8d docs. 218 tests passing, ruff clean. Known
  limitations at the time (documented in README): folder scan skips
  already-imported AOIs, export covers template AOIs only, a template
  workbook is still required, and column types/descriptions fall back to
  untyped if the source .L5X leaves the import folder. The "no in-place
  update" limitation was resolved by Phase 9c (Nov 24, 2026): a re-import
  now shows a diff and updates the AOI in place. User-defined AOIs (Erik's
  own plant AOIs, not in the PlantPAx library) appear as tabs alongside the 56.
  - Input: Studio 5000 `.L5X` exports. Folder scan — a configured directory
    (default `aoi_imports/`, gitignored) scanned on startup plus a rescan
    endpoint/button — and single-file upload in the UI.
  - Parser (`src/pypax/l5x.py`, ElementTree only, never executes anything):
    AOI name, TargetRevision, parameters + local tags (name, DataType,
    Description, DefaultData decoded per type: BOOL 1 byte, DINT/REAL 4-byte
    little-endian, STRING length-prefixed ASCII). Nested sub-objects are
    discovered mechanically: LocalTags whose DataType is another AOI defined
    in the same file (e.g. `FullStall` → P_Alarm, `CmdSrc` → P_CmdSrc) expand
    into dotted columns. The real P_ValveMO export stays in
    `workspace/user/files/` and is NEVER committed to the public repo; tests
    use a small synthetic L5X fixture built in-test.
  - Curation rule (Erik, Nov 24, 2026; verified against
    P_ValveMO_4.10.06_AOI.L5X, where all 82 sheet columns obey it): a column
    is a config value = its leaf name starts with `Cfg_`. Top-level
    `Cfg_{Alarm}AckReqd|ResetReqd|Severity` and `Cfg_Has{Alarm}Alm`
    pass-throughs are dropped in favor of the nested
    `{Alarm}.Cfg_AckReqd|Cfg_ResetReqd|Cfg_Severity|Cfg_Exists`. Order:
    Cfg_Desc/Label/Tag/Area first, then the rest in L5X declaration order.
    Everything else (Inp_/Out_/Sts_/PCmd_/XCmd_/Wrk_/HMI_/...) is excluded.
  - Propose + confirm: import produces a proposal (AOI name, revision, column
    list with types/descriptions/defaults, nested objects found) shown in the
    UI with checkboxes — e.g. P_ValveMO proposes IntlkTrip, Erik unchecks it,
    matching the sheet. Until confirmed, no tab appears. Re-import of the same
    AOI shows a diff (added/removed/renamed columns) and needs a fresh confirm.
  - Storage: `config/aoi/custom/<AOI>.yaml` (same shape as `v4.10.6.yaml`,
    plus a header with source L5X name, revision, import date). The backend
    AOI model is template sheets ∪ custom YAML. A name collision with an
    existing AOI (PlantPAx or custom) is rejected with a clear error unless
    it is a re-import diff.
  - Safety: custom AOIs go through the exact same write path (PLC-name check,
    typed confirmation, chunked writes, per-tag read-back verification). A
    custom AOI whose members don't exist on the PLC fails read-back per tag
    (verified / mismatch / failed) — never silently. Nothing is writable from
    a proposal; only confirmed AOIs are.
  - Frontend: tabs already come from `/api/aois`, so no tab changes — just the
    import UI (upload + scan results + proposal review + confirm) and a
    "custom" badge on imported tabs.
  - Audit test: custom AOIs are allowed without a template sheet; the test
    asserts every custom YAML entry's columns all carry `Cfg_` leaves and the
    YAML parses. Docs: README section for import/scan.
  - Out of scope: hand-editing an imported AOI's columns in the UI (re-import
    instead); writing `.L5X` back out.
  - After Phase 8 lands, a dedicated audit pass re-checks the whole codebase
    for missed errors/bugs (write safety, UI bugs, other issues) and fixes
    what it finds — Erik's standing instruction, Nov 24, 2026: be SUPER
    skeptical about everything. Adversarial review: assume every safety
    path, every UI flow, and every assumption is wrong until proven
    otherwise; verify claims against the running code, not the reports.
    DONE Nov 24, 2026 (merged 33d7860, 233 tests): 11 findings fixed,
    worst first — (1) a write silently reverted unstaged PLC values that
    had drifted since the tab loaded (write now diffs only staged tags);
    (2) blank line inside a paste shifted rows onto the wrong instance;
    (3) a newline in a scan-dir filename injected a second AOI into the
    custom YAML; (4) folder scan had no 5 MB cap; (5) AOI collisions were
    case-sensitive; (6) PYPAX_HOST=0.0.0.0 bound the unauthenticated API
    to all interfaces (loopback-only now); (7) connect was broken with no
    profile (UI sent "default"); (8) dev mock stale vs contract (export
    501, paste staged shape, confirm 404); (9) import/remove had no audit
    trail; (10) lock-free /api/changes + /api/audit could 500; (11) paste
    into untyped BOOL columns wrote True on a real PLC. XXE/billion-laughs
    verified safe (expat doesn't expand internal entities).

## 7. Testing (no live PLC available)

- pytest: pure-function unit tests (paste/TSV parsing, diffing, BOOL/REAL/INT/STRING
  coercion incl. the REAL `.6f` rounding and non-string write cases, AOI detection,
  array-dimension expansion, PLC-name filename guard).
- Integration tests against the fake PLC: full read → edit → diff → write → read-back
  cycle, chunking, failure injection (bad tag, mismatch on read-back, wrong PLC name).
- Regression: CLI golden tests + spreadsheet round-trip tests on every phase.
- Live verification is manual, by Erik, using the Phase 7 checklist — Muse cannot
  do it from here.

## 9. Execution: git worktrees + parallel agents

Yes — work in git worktrees so parallel agents never step on each other.

- Integration checkout: `~/workspace/pypaxtagtool` on `feature/webapp`.
- One worktree per phase/chunk, outside the repo dir, e.g.
  `~/workspace/pypaxtagtool-worktrees/phase-1-core/`, each on its own branch
  (`feature/webapp-phase-1-core`, …). One subagent per worktree, one phase at a time.
- Frozen contracts before fanning out: core library API, fake-PLC interface, REST
  API shape, config schema — so parallel agents can't drift apart.
- Sequencing: Phase 0 + 1 are foundation work, done first by a single agent
  (shared core files). Then fan out in parallel: Phase 2 (PLC optimization),
  Phase 3 (backend), Phase 4 (frontend). Phases 5–6 depend on 3–4 and integrate after.
- Each agent commits small chunks on its own branch and must pass its tests locally
  before handing back. The main agent integrates: reviews each diff, resolves
  conflicts, merges into `feature/webapp` (per-phase PR into `feature/webapp` when
  a reviewable diff is useful), runs the full suite + cross-platform CI.
- Worktrees/branches are deleted after merge. No two agents ever push to the same branch.

## 10. Phase 9 — Roadmap features 1–3 (approved Nov 24, 2026)

DONE Nov 24, 2026 — 9a snapshots (f96466b7), 9b string-UDT learning
(d103759d), 9c re-import in place (2757752c), integrated into
`feature/webapp` (bb7b2b9, 0cb0ded, 4d9358d), then a skeptic audit pass
(40b91d8) fixed 8 findings — worst was the auto-snapshot capturing the
stale tab baseline (same failure mode as the Phase 8 audit's worst
finding). 311 tests passing, ruff clean. Local only; push to GitHub
happens all at once when Erik says so.

Roadmap items 1–3 from `ROADMAP.md`, built in parallel by three agents
(9a snapshots, 9b string-UDT learning, 9c re-import in place). Rules
for this phase, from Erik: **local commits only — never prompt to push
to GitHub, never push during the work; when the work + skeptic audit
are good, push everything at once at the end.** After integration, a
dedicated skeptic audit pass runs before anything is called good.

### 9a. PLC snapshots (`src/pypax/snapshots.py`, new)

- Snapshot file: JSON, format `"pypax-snapshot/1"`:
  `{format, plc_name, created_at (ISO 8601 with offset), aoi_version,
  aois: {AOI: {instance: {subtag: value}}}}`. Values are JSON-native
  (REAL float, BOOL 0/1 int, strings str). Covers template + custom
  AOIs.
- Storage: `snapshots/` (gitignored), auto-snapshots under
  `snapshots/auto/`; filename `<plc>_<YYYYMMDD-HHMMSS>.json`;
  manual names sanitized to `[A-Za-z0-9_-]+` (no path separators).
  A manual name must not collide with an auto-snapshot stem (and an
  auto save suffixes instead of shadowing a manual one) — lookup by
  name is single-valued, so a shared stem would leave one snapshot
  unreachable; the listing dedupes the same way (manual wins).
- Endpoints (frozen):
  `POST /api/snapshots` `{name?}` → 201 `{name, plc_name, created_at,
  aoi_count, instance_count}`; `GET /api/snapshots` → list newest
  first with `auto` flag; `GET /api/snapshots/{name}` → download
  (raw bytes, even for a corrupt file — it must stay inspectable and
  deletable); `POST /api/snapshots/{name}/compare` → `{plc_name,
  differences: [{aoi, instance, subtag, snapshot_value,
  current_value}]}` (409 when not connected, 404 unknown name);
  `POST /api/snapshots/upload`
  (multipart `file`, `.json` only, 5 MB cap) → 409 when the file's
  `plc_name` differs from the connected PLC; `POST
  /api/snapshots/{name}/restore` → stages values that differ from the
  current baseline through the normal staging path (validated via
  `coerce_value` with the app's configured string registry; failures
  reported as `skipped` with reasons, never staged) and returns
  `{staged, skipped}` — **restore stages only, it never writes**;
  `DELETE /api/snapshots/{name}` (removes the file without parsing
  it, so a corrupt snapshot is deletable).
- Auto-snapshot before every write: `POST /api/write` captures the
  pre-write baseline to `snapshots/auto/` first — freshly re-read
  from the PLC, not the cached tab baseline, so a value that drifted
  since the tab loaded is captured as it was before the write;
  keeps the newest 10 per PLC, prunes older. A failed auto-snapshot
  never blocks the write — it is logged, recorded in the audit entry,
  and reported as `auto_snapshot: {ok: false, error}` in the write
  response. Manual snapshots are likewise freshly re-read.
- Frontend: Snapshots panel (save with name prompt, list, compare
  diff table, restore, download, delete). Dev mock server gets
  in-memory snapshot stubs.
- Tests: unit tests (build/diff/save/load, name sanitization, corrupt
  JSON rejection, oversize upload) + API tests (connect fake PLC →
  snapshot → move the fake's values → compare shows the diff →
  restore stages → write → read-back verified).

### 9b. Learn string-UDT sizes from L5X imports

- `l5x.py`: new `learn_string_sizes(root) -> dict[str, int]` — scans
  `<DataType Family="StringFamily">` blocks, reads each `DATA` SINT
  member's `Dimension` → size, keyed by upper-cased type name. Only
  StringFamily UDTs with a SINT DATA member count; anything else is
  ignored, never guessed.
- Registry: new `src/pypax/string_sizes.py` (pinned sizes + learned
  sizes, merged). Learned sizes persist in
  `config/string_udt_sizes.yaml` — `{STRING_COND: {size: 20, source:
  "P_Intlk_4.10.06_AOI.L5X"}}` — committed (type names, sizes, source
  filenames only; no PLC/site details). The import confirm path
  merges learned sizes in — on the fresh-import path AND the
  re-import (update) path; a learned size that conflicts with a pinned
  size is ignored and surfaced as a warning in the import response.
- `paste.py`: `_string_limit` consults the merged registry (pinned
  first, then learned); the pinned map stays the fallback when the
  registry file is absent.
- Frontend: new `GET /api/string-limits` → merged map; `app.js`
  fetches it at startup and uses it, falling back to its pinned map
  until loaded.
- Tests: parser tests on a synthetic L5X (several StringFamily sizes,
  a non-string DataType ignored, a StringFamily UDT with no DATA
  member ignored), registry merge/conflict tests, paste tests with
  learned sizes via a tmp registry file, API test for
  `/api/string-limits`.

### 9c. Re-import a custom AOI in place

- `POST /api/aois/import` with an already-imported AOI name
  (case-insensitive) no longer returns 409. It returns `{status:
  "update_available", proposal_id, diff: {added: [columns], removed:
  [columns], changed: [{subtag, before, after}], unchanged: n,
  before_details_known: bool}}`;
  a column is "changed" when type, description, or default differs.
  `before_details_known` is false only for a legacy import whose
  confirmed details survive nowhere — then only added/removed are
  detectable and the UI says so.
- `POST /api/aois/import/confirm` accepts an update proposal and
  rewrites `config/aoi/custom/<AOI>.yaml` in place. It refuses with
  409 when staged changes exist on that AOI (a re-import must not
  silently drop staged edits — discard or write first). Confirming
  also writes `config/aoi/custom/<AOI>.meta.json` — the confirmed
  column details (type/description/default) — so the diff's "before"
  side survives restarts and a deleted source `.L5X`; deleting the
  custom AOI removes the sidecar with the YAML.
- An upload whose filename already holds a *different* AOI's export
  is refused (400) — the upload is persisted before confirm, so
  overwriting it would destroy that export. Same AOI, same filename
  is a re-import and is allowed.
- Scan: `GET /api/aois/scan` also reports already-imported AOIs whose
  source file changed (`update_available: true`, `source_file`).
  Detection is a SHA-256 hash of the source file (5 MB cap), stored
  in the YAML metadata at confirm time.
- Frontend: the scan list shows "Update available" on already-imported
  AOIs; confirming shows the diff (added green, removed red, changed
  old → new) with an "Update AOI" button. No column checkboxes on
  re-import — it is all-or-nothing with a visible diff.
- Tests: re-import of an unchanged file → empty diff; changed file →
  correct diff; confirm updates the YAML; confirm with staged changes
  → 409; scan detects a changed file; case-insensitive name match.

### File ownership for Phase 9 (avoid merge conflicts)

- 9a owns: `src/pypax/snapshots.py`, snapshot routes + auto-snapshot
  hook in `app.py`, snapshots markup in `index.html`, snapshot UI in
  `app.js` (new functions only — do not touch `stringLimit` /
  `parseValue`), dev mock snapshot stubs, `tests/test_snapshots.py`,
  `.gitignore` (`snapshots/`), README snapshots section.
- 9b owns: `learn_string_sizes` in `l5x.py`,
  `src/pypax/string_sizes.py`, `_string_limit` in `paste.py`,
  `GET /api/string-limits` + learned-size merge in the import confirm
  path in `app.py`, string-limits fetch in `app.js` (localized to that
  area), `config/string_udt_sizes.yaml`, `tests/test_string_sizes.py`,
  `tests/test_web_paste.py` additions, README validation section.
- 9c owns: update path + hash storage in `custom_aoi.py`, import /
  confirm / scan changes in `app.py`, `tests/test_custom_aoi_import.py`
  (new file — do not edit shared test files), README Custom AOIs
  section.
- Shared files (`app.py`, `app.js`) are edited in different regions by
  design; the integrating agent resolves any overlap. All three agents:
  local commits only, no GitHub pushes, small commits, full suite +
  ruff green in the worktree before hand-back, synthetic fixtures only
  (never open or commit Erik's real `.xlsm` / `.L5X` files), `pathlib`
  everywhere.

## 8. Open questions (do not block starting)

- Where it runs: Erik's laptop on the plant network? Windows or Mac? (Decides packaging.)
- AG Grid for the grid — OK, or a preference?
- First live smoke test: which PLC / which harmless tag? (Decided at handover, not now.)
