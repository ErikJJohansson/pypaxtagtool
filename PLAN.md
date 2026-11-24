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

## 8. Open questions (do not block starting)

- Where it runs: Erik's laptop on the plant network? Windows or Mac? (Decides packaging.)
- AG Grid for the grid — OK, or a preference?
- First live smoke test: which PLC / which harmless tag? (Decided at handover, not now.)
