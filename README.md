# PyPaxTagTool — PlantPAX Tag Configuration (CLI + Webapp)

Read and write PlantPAX AOI configuration tags on a PLC over EtherNet/IP,
using Python instead of Excel VBA. Two interfaces share one core library,
so they cannot drift apart:

- **Webapp** — a browser grid (one tab per AOI) for reading, editing, and
  writing config tags, with copy-paste from spreadsheets, changed-cell
  highlighting, and read-back-verified writes.
- **Legacy CLI** — `pypax.py`, the original read/write tool that works
  directly with the PlantPAX configuration spreadsheet.

Target: PlantPAX library **v4.10.6**. Template:
`ProcessLibraryOnlineConfigTool.xlsm` only (`SpotConfigTool.xlsm` is out of
scope — it is ignored entirely).

## Why Python instead of VBA?

The original PlantPAX tag configuration tool runs in Excel using VBA, and
the computer running the macro also needs RSLinx Classic installed.

The Python version drops both requirements: no Excel and no RSLinx on the
machine, and it runs on macOS and Windows alike (pure Python — no COM,
no RSLinx, no Windows-only dependencies). It also finds the tag instances
for you: the original tool requires you to supply every instance name, while
this one asks the PLC how many instances of each AOI exist, so you won't
miss any.

## Requirements

- Python **3.11+** (see `requires-python` in `pyproject.toml`)
- Network access to the PLC (EtherNet/IP)
- The PlantPAX template workbook `ProcessLibraryOnlineConfigTool.xlsm`
  (in the working directory, or pointed to by config — see
  [Configuration](#configuration))
- macOS, Windows, or Linux — the stack (pycomm3, openpyxl, FastAPI/uvicorn)
  is pure Python and cross-platform. `pathlib` is used everywhere; there
  are no hardcoded drive letters or backslashes.

## Install

Clone the repo, create a virtual environment, and install in editable mode
with the dev tools (verified on this machine, Python 3.12):

```bash
git clone https://github.com/ErikJJohansson/pypaxtagtool.git
cd pypaxtagtool
git checkout feature/webapp        # the webapp lives on this branch for now
python3 -m venv .venv
.venv/bin/pip install -e ".[dev]"
```

On Windows, use `.venv\Scripts\pip install -e ".[dev]"` and
`.venv\Scripts\pypax-web` below. (The original, unquoted
`pip3 install requirements.txt` line from the old README is wrong —
`requirements.txt` is a file to install *from*: `pip install -r requirements.txt`.
Prefer the `pip install -e ".[dev]"` form above; it installs the package
plus pytest and ruff.)

## Launching the webapp

```bash
.venv/bin/pypax-web
```

Then open <http://127.0.0.1:8000/> in a browser. The server binds to
`127.0.0.1` only (localhost) — it is never exposed on the network, and
there is no multi-user mode or auth in v1.

> **Launch gotcha (verified):** `python -m pypax.web` does **not** work
> from the repo root — the legacy `pypax.py` shim at the root shadows the
> installed `pypax` package and Python fails with
> `ModuleNotFoundError: __path__ attribute not found on 'pypax'`.
> Use the `pypax-web` command. (The module form works from any other
> directory.)

## Connecting to a PLC

In the header of the webapp, enter a **commpath** (PLC IP address and slot,
e.g. `10.10.17.10/4`; without a slot it defaults to slot 0) and optionally a
**profile** name, then click **Connect**. The banner shows the connection
state, the PLC name read back from the PLC (`get_plc_name()`), and the comm
path.

Profiles come from a config file (see [Configuration](#configuration)).
Each profile holds a comm path, an expected PLC name, an AOI version, and a
safety level:

- `test` — writes freely (still subject to the PLC-name double-check below).
- `production` — requires the name double-check **and** the typed
  confirmation in the write dialog. The app refuses to write if the
  profile's expected PLC name does not match the PLC it actually
  connected to.

The PLC-name double-check runs twice: once when the write is submitted,
and once against the profile's expected name. A mismatch is a `409`
response and nothing is written. If a connect attempt fails, the UI reports
the error and stays disconnected.

## Using the grid

- **One tab per AOI** (56 in the v4.10.6 library, plus any custom AOIs
  you import — see "Custom AOIs (L5X import)"), loaded lazily — a tab's
  data is fetched from the PLC the first time you open it. Columns are the
  AOI's subtags (header shows the subtag and its data type); rows are the
  tag instances, with the tag name pinned in the first column (read-only).
- **Paste from a spreadsheet:** copy cells in Excel and paste them into
  the grid. The top-left cell of the pasted block lands on the cell you
  pasted into (the *anchor*); values flow right across subtags and down
  across instances. Every value is checked against the column's data type
  (from row 6 of the template). Bad values are flagged **red** with a
  reason and are never staged — the cell reverts to the PLC value.
- **BOOL cells are checkboxes**, not 0/1 text: a ticked box is 1, an
  empty box is 0. Single-click a BOOL cell to open its checkbox editor
  and toggle it; typing 0/1/true/false still works too.
- **String lengths are checked like the spreadsheet flags them** —
  see "Data validation" below for the per-type limits.
- **Changed cells** — any cell whose value differs from the last-read PLC
  baseline is highlighted **amber**. The staged-changes panel lists every
  staged change (AOI, instance, subtag, old → new value) and is the source
  of truth for what a write will do.
- **Discard:** `DELETE /api/changes` clears all staged changes on the
  server (baseline untouched). The **Discard staged changes** button in
  the UI calls it, then clears the local amber/red highlighting.
- **Audit:** the audit view lists every connect, disconnect, and write,
  with per-tag results. Entries are also appended to
  `audit/audit-YYYY-MM-DD.jsonl` in the working directory.

Data in the grid comes from the PLC itself, not from the spreadsheet:
the subtag columns come from the template workbook, and the values are
the live PLC baseline.

## Data validation

Every value is checked against the column's data type (row 6 of the
template sheet) **before** it is staged — twice, in two places that
agree: the grid's `parseValue` in `src/pypax/web/static/app.js` (for
instant feedback while typing or pasting) and the server's
`coerce_value` in `src/pypax/web/paste.py` (the authority — direct
cell edits are staged through the paste endpoint, so they get the same
check as a pasted block). A value that fails is flagged **red** with a
reason, the cell reverts to the PLC baseline, and the value is never
staged or written.

| Data type | Accepted values | Rejected |
| --- | --- | --- |
| BOOL | `0`, `1`, `true`, `false` (also `yes`/`no`/`on`/`off` on paste) | anything else — e.g. `2`, `maybe` |
| SINT / INT / DINT / LINT | whole numbers, optional leading `-` | decimals, text, empty |
| REAL | any finite number (`90`, `3.5`, `-0.25`) | text, `nan`, `inf` |
| STRING_* | text up to the type's character limit (below) | text longer than the limit |

String limits, in characters — the same limits the spreadsheet flags:

| Type | Limit | Source |
| --- | --- | --- |
| STRING_N (STRING_8, STRING_12, …) | N | UDT definition: DATA SINT[N] |
| STRING (the Logix type) | 82 | Logix STRING: DATA SINT[82] |
| STRING_Area | 8 | P_ValveMO 4.10.06 export: DATA SINT[8] |
| STRING_Cond | 20 | P_Intlk 4.10.06 export: DATA SINT[20] |
| STRING_NavTag | 20 | P_Intlk 4.10.06 export: DATA SINT[20] |

A longer string is rejected with its limit ("string is 21 characters,
longer than the 20-character limit for STRING_Cond"). The check exists
because the PLC would not hold the extra characters either — pycomm3's
string encoder does no truncation, so without it an over-long string
would reach the write call unchecked.

**Learning new string sizes.** A custom string UDT is not limited to
the table above: when a `.L5X` is imported, the app reads the
`<DataType Family="StringFamily">` blocks of that file and learns
each type's size from its `DATA` SINT dimension (e.g. a `STRING_Foo`
with `DATA SINT[33]` holds 33 characters). Learned sizes persist in
`config/string_udt_sizes.yaml` as `{NAME: {size, source}}` and apply
in both the grid and the server (`GET /api/string-limits` serves the
merged map). Teaching a new size is just importing an `.L5X` that
defines the type — no code change. Two guardrails: a learned size
that conflicts with a pinned size above (or with the `STRING_N`
naming rule, or with an earlier learned size for the same type) is
ignored and reported as a warning in the import response, and a type
with no known size is kept as text with no length check.

Two special cases:

- **Untyped columns accept nothing.** A custom AOI whose source
  `.L5X` was deleted from the import folder loses its column types
  (types come from the cached proposal metadata). With no type the
  value cannot be validated — and pycomm3 encodes any truthy value as
  BOOL true — so every paste into such a column is rejected until the
  `.L5X` is re-imported.
- **Types the rules don't know** (anything not in the tables above)
  are kept as text with no length check. If a new string UDT shows up,
  pin its size in `coerce_value` / `stringLimit` (with the export it
  came from, as above) and add a test in `tests/test_web_paste.py`.

## Writing to the PLC

1. Edit cells or paste a block. Amber cells are staged.
2. Click **Write**. A confirmation dialog shows the connected PLC name
   and the number of staged changes per AOI.
3. **Type the PLC name exactly** to confirm — the confirm button stays
   disabled until it matches. This is the second half of the double-check.
4. The app writes only the staged differences, in chunks of at most 100
   tags per PLC request, then **re-reads every written tag** and reports
   each one as:
   - `verified` — the PLC now holds the intended value,
   - `mismatch` — the read-back disagrees (the tag stays staged so you
     can review or retry it),
   - `failed` — the write itself failed (also stays staged).
5. Verified tags are unstaged and become the new baseline. Mismatches and
   failures are never silently dropped — they stay in the staged panel.

If a write reports mismatches, do **not** re-write blindly: check the tag
in Studio 5000, check the audit log for what was actually sent, and see
`docs/SMOKE_TEST.md` for the mismatch procedure.

Writes are also protected by a single-connection lock: one PLC connection
at a time, no parallel writes to the same PLC.

## PLC snapshots

A snapshot is the PLC baseline for every AOI (template + custom) at
one moment, saved as a JSON file
(`pypax-snapshot/1`: PLC name, timestamp, AOI version, all values).
Snapshots live in `snapshots/` (gitignored; override with
`PYPAX_SNAPSHOT_DIR`).

- **Save:** the **Snapshots** panel's **Save snapshot…** asks for a
  name (`[A-Za-z0-9_-]+`) and stores it; without a name it is saved
  as `<plc>_<YYYYMMDD-HHMMSS>.json`. A name that already exists is
  refused — a named snapshot is never silently overwritten.
- **Auto-snapshot before every write:** each write first captures
  the pre-write baseline to `snapshots/auto/` — freshly re-read from
  the PLC, so a value that drifted since a tab was loaded is captured
  as it was before the write — so any write can be rolled back from
  its own snapshot. The newest 10 auto-snapshots per PLC are kept;
  older ones are pruned. A failed auto-snapshot never blocks the
  write — it is logged, recorded in the audit entry, and reported in
  the write response.
- **Compare:** pick a snapshot and click **Compare** — the diff table
  shows every tag that differs from the live PLC (snapshot value vs
  current value), including tags that vanished or appeared since the
  snapshot.
- **Restore:** **Restore** stages the snapshot's differing values
  through the same validation as a paste (bad values are skipped with
  a reason, never staged). Restore only stages — nothing is written
  until you go through the normal write confirmation. A snapshot for
  another PLC is refused.
- **Download / delete:** download any snapshot as its `.json` file;
  a snapshot file from another machine can be uploaded back via
  `POST /api/snapshots/upload` (`.json` only, 5 MB cap, PLC name must
  match the connected PLC).

## Export to spreadsheet

**Export .xlsx** downloads `<plc_name>_ConfigTags.xlsx` in the exact template layout (headers rows 4–9, instances from row 10, Setup sheet counts), with baseline values from the PLC and staged changes applied on top, highlighted amber. The exported file works with the legacy CLI below and vice versa — that round-trip is the point. When disconnected, the template is exported as-is under the fallback name `PyPax_ConfigTags.xlsx`. Export covers the template AOIs only. Custom AOIs are not exported —
a custom AOI requires its logic (in Studio 5000), and the export
would only carry its data, so the spreadsheet round-trip doesn't
apply to them.

## Custom AOIs (L5X import)

The webapp can learn AOIs the template doesn't have, straight from a
Studio 5000 `.L5X` export. It never executes the file — the export is
parsed as XML.

- **Two ways in:** click **Import AOI** and pick a `.L5X` file (5 MB
  cap), or drop `.L5X` files into the `aoi_imports/` folder (override with
  `PYPAX_AOI_IMPORT_DIR`) and click **Scan folder**. Scanned files are
  proposed, never auto-confirmed.
- **Propose, then confirm:** the proposal shows the AOI name, its
  revision, the nested sub-objects it found (e.g. alarm members under
  `FullStall`), and every candidate column with its type, description,
  and default — each with a checkbox. Only columns you tick are imported.
  Which columns are candidates is decided by the curation rule: a column
  is a config value when its leaf name starts with `Cfg_` (top-level
  alarm pass-throughs like `Cfg_FullStallSeverity` map to the nested
  member `FullStall.Cfg_Severity`). Non-`Cfg_` members (inputs, outputs,
  status) are never offered.
- **After confirm** the AOI gets its own tab with a `custom` badge and
  behaves exactly like a template AOI: same grid, same paste validation,
  same write path (PLC-name check, typed confirmation, chunked writes,
  per-tag read-back verification). Its definition is stored in
  `config/aoi/custom/<AOI>.yaml`.
- **Re-import in place:** to update an imported AOI, upload the new
  `.L5X` export (or drop it into the import folder and click **Scan
  folder**) — the app answers with a diff instead of a fresh proposal:
  columns added (green), removed (red), and changed (old → new for
  type, description, or default). **Update AOI** rewrites
  `config/aoi/custom/<AOI>.yaml` in place; it is all-or-nothing, with
  no column checkboxes. The folder scan spots a changed export by
  comparing the SHA-256 stored in the YAML header at confirm time
  against the file on disk, and shows "Update available" on it. An
  update is refused (409) while staged changes exist on that AOI —
  discard or write them first, so a re-import never silently drops
  an edit. The confirmed column details (type, description, default)
  are also stored next to the YAML in
  `config/aoi/custom/<AOI>.meta.json`, so the diff still works after
  a restart or after the source `.L5X` is deleted; for an AOI imported
  before that file existed the diff can only show added/removed
  columns and says so.
- **Remove:** the toolbar's **Remove custom AOI** button deletes a custom
  AOI (template AOIs can't be removed).
- **Limits:** a template workbook is still required (the app won't serve
  custom AOIs without one); column types and descriptions come from the
  source `.L5X` kept in the import folder, so if that file is deleted
  the columns become untyped and pastes into them are rejected until
  the `.L5X` is re-imported (the app cannot validate a value it cannot
  type, and an unvalidated BOOL write would land as true). Real
  Rockwell exports are test fixtures only and must never be committed
  — tests use synthetic `.L5X` files.

## The legacy CLI

The original command-line tool still works, unchanged, and shares the core
library with the webapp. `pypax.py` at the repo root is a thin shim over
`src/pypax/cli.py`.

```
pypax.py 10.10.17.10/4 read [ProcessLibraryOnlineConfigTool.xlsm]
pypax.py 10.10.17.10/4 write PLCNAME_ConfigTags.xlsx
```

`10.10.17.10/4` is the PLC IP address and slot number. Without a slot
number (just the IP) it defaults to slot 0. If no file is given for `read`,
the template `ProcessLibraryOnlineConfigTool.xlsm` in the repo is used.

**Read mode.** A new file is created from the template, named
`PLCNAME_ConfigTags.xlsx`. The script loops through each sheet that looks
like a PlantPAX AOI, asks the PLC for the number of instances of each
type (written to the `Setup` sheet), then bulk-reads every instance's tags
and writes them into the AOI sheet. Edit the values you want to change in
the output file.

**Write mode.** Uses the file you pass. Only the changes detected between
the PLC tag values and the spreadsheet values are written, in chunks of at
most 100 tags. If the filename does **not** contain the name of the PLC
being written to, the script exits — this is to ensure you are writing to
the correct PLC. The number of instances written is determined by the rows
in each AOI sheet: there must be no spaces between names in the names
column, otherwise the blank is treated as the last row to write.

The template `.xlsm` in the repo holds every PlantPAX AOI type and its
configuration tags. Be careful modifying it — preferably, don't.

**Troubleshooting (CLI).** Can you ping the PLC? Make sure you have network
connectivity before running the script. `ping 10.10.17.10` works the same
on macOS/Unix and Windows.

## Configuration

Three layers, one source of truth each. Real values are never committed —
this repo is public.

- **Spreadsheet layout** (rows/columns, Setup columns): `src/pypax/layout.py`,
  read by the CLI, webapp, and tests.
- **AOI definitions:** versioned data files — `config/aoi/v4.10.6.yaml`,
  generated from the template spreadsheet by
  `scripts/extract_aoi_definitions.py`. A test asserts the checked-in file
  matches the spreadsheet. `AOI_definitions.py` stays as a compatibility
  shim (still the runtime source for the CLI).
- **App / connection policy:** copy `config/config.example.toml` to
  `config.toml` (gitignored) or to a per-PLC profile file under
  `config/plcs/<name>.toml` (also gitignored). It holds the template path,
  host/port, write chunk size, REAL tolerance, read-back retries, and
  `[[profiles]]` entries (commpath, expected PLC name, AOI version,
  safety level).

Settings precedence: **`PYPAX_*` environment variables** beat the TOML
`[app]` section, which beats the built-in defaults. The config file is
only read when `PYPAX_CONFIG_FILE` points at it. Unknown `[app]` keys fail
fast at startup. Defaults: host `127.0.0.1`, port `8000`, write chunk size
`100`, REAL tolerance `1e-6`, read-back retries `1`.

```bash
# Example: point at a config file and override the port
PYPAX_CONFIG_FILE=config.toml PYPAX_PORT=8001 pypax-web
```

## Development

```bash
.venv/bin/python -m pytest -q        # full suite (218 tests after Phase 8)
.venv/bin/ruff check src tests scripts
.venv/bin/ruff format --check src tests scripts
```

- **Fake PLC:** `tests/fakes/fake_plc.py` is an in-memory `LogixDriver`
  stand-in with the same interface (`tags`, `read`, `write`,
  `get_plc_name`). All tests run against it — no live PLC is available
  here, and pycomm3 itself is already proven. The fake also has
  test-only hooks (`corrupt_after_write`, `drop_after_write`) to simulate
  read-back mismatches.
- **Test fixtures:** small workbooks are generated in code
  (`tests/helpers.build_workbook`); the ~4 MB template `.xlsm` is not
  copied around.
- **Workflow:** one git worktree per phase/chunk, one branch per worktree,
  small commits, tests green before hand-back. The integrator merges into
  `feature/webapp` and pushes via the GitHub API. **Never merge to `main`
  without Erik's explicit say-so** — `main` is the release branch.
- **CI:** `.github/workflows/ci.yml` (macOS + Windows + Linux matrix)
  exists only in the local branch. The GitHub App used to push has no
  Workflows permission, so the file cannot be pushed via the API and
  GitHub Actions will not run until it is added manually on GitHub.
  Until then, local `pytest` + `ruff` are the gate.

## Known limitations

- **No live PLC has been tested against yet.** Everything is fake-PLC
  tested. First live contact must follow `docs/SMOKE_TEST.md`
  (read-only first, then one harmless single-tag write).
- **`P_ValveMO` spelling (resolved Nov 24, 2026):** the template
  sheet's `.Cfg_CoastToLS` was confirmed correct against Erik's
  Studio 5000 `P_ValveMO_4.10.06_AOI.L5X` export; `AOI_definitions.py`
  had the typo (`.Cfg_CoastTolS`) and was fixed. All 82 P_ValveMO
  subtags now match across spreadsheet, definitions, and YAML.
- **SpotConfigTool.xlsm is out of scope** and ignored.
- **Custom AOIs:** re-import updates in place (see "Custom AOIs (L5X
  import)"); the folder scan skips already-imported AOIs whose file
  is unchanged, and export covers template AOIs only.
- **No lock file yet** — dependencies are exact pins in `pyproject.toml`
  / `requirements.txt`.
- The `python -m pypax.web` repo-root launch caveat above.
