# First live PLC smoke test — checklist

Follow this on the plant floor, in order. Everything before this point was
tested against the fake PLC only — this is the first time the tool talks to
real hardware. Do the read-only part first; do not skip ahead to writing.

Prereqs: the webapp installed (`pip install -e ".[dev]"`), the template
`ProcessLibraryOnlineConfigTool.xlsm` in the working directory, and a
**test PLC** if one exists — otherwise the least critical production PLC
you can get a maintenance window on. Have Studio 5000 open on the same
machine (or a second screen) so you can cross-check values independently.
Have this file open — tick each box as you go.

## 0. Before you connect

- [ ] You are on the plant network and can ping the PLC:
      `ping <plc-ip>` replies.
- [ ] You know the PLC's name as shown in Studio 5000 (controller
      properties). Write it here: `PLC name: ______________________`
- [ ] You know the comm path: `Commpath: ______________________`
      (IP/slot, e.g. `10.10.17.10/4`; slot 0 if no slot is given.)
- [ ] The template workbook is present and unmodified.
- [ ] You have a backup of the current config: if a recent
      `<PLC>_ConfigTags.xlsx` exists, keep it; if not, the read-only step
      below produces one.

## 1. Read-only smoke test

- [ ] Launch the webapp: `pypax-web`, open <http://127.0.0.1:8000/>.
- [ ] Enter the commpath and click **Connect**.
- [ ] The banner shows the PLC name — confirm it matches the name you
      wrote down in step 0. If it does not match, **stop**: you are talking
      to the wrong PLC.
- [ ] Open one small AOI tab, e.g. **P_DIn**. Wait for it to load.
- [ ] Pick 3–5 tags you can check independently. For each, compare the
      grid value against Studio 5000 (online tag monitor) or the last
      known-good spreadsheet:

      | Tag | Grid value | Studio 5000 / spreadsheet value | Match? |
      |-----|------------|----------------------------------|--------|
      |     |            |                                  |        |
      |     |            |                                  |        |
      |     |            |                                  |        |

- [ ] If every value matches: the read path works. If any value does not
      match, stop and note it — do not proceed to writing.
- [ ] Click **Disconnect**. Check `audit/audit-YYYY-MM-DD.jsonl` in the
      working directory: it should contain a `connect` entry and a
      `disconnect` entry.

Optional cross-check with the CLI (same read path, different interface):

```bash
python pypax.py <ip/slot> read
```

This writes `<PLC>_ConfigTags.xlsx`. Spot-check a few values in it
against Studio 5000 the same way.

## 2. One harmless single-tag write

Pick a tag that is safe to change and easy to change back. Good candidates:
a `.Cfg_Desc` string on an unused instance, or a `.Cfg_HiLim` on an
analog input that is out of service. Do **not** pick an interlock, a
permissive, a running motor command, or anything in a running loop.

- [ ] Tag chosen: `Tag: ______________________`
      Old value: `______________________`  New value: `______________________`
- [ ] Reconnect in the webapp and open the tag's AOI tab.
- [ ] Edit the one cell (or paste the one value). Confirm the cell turns
      amber and the staged-changes panel shows exactly one change.
- [ ] Click **Write**. The dialog shows the PLC name and 1 staged change.
- [ ] **Type the PLC name exactly** and confirm. (The button stays
      disabled until it matches.)
- [ ] The results panel reports the tag as `verified`.
- [ ] Confirm the new value in Studio 5000 (online tag monitor).
- [ ] Write the tag back to its old value the same way, and confirm it
      reports `verified` again. The PLC is now exactly as you found it.

## 3. P_ValveMO spelling check

There is one known discrepancy between the template spreadsheet and
`AOI_definitions.py`, recorded as a test exception because neither file
has been touched: **P_ValveMO position 10**.

- [ ] In Studio 5000, open a real P_ValveMO instance and look at the
      member at position 10. Record the spelling:
      `Real spelling: ______________________`
      (`.Cfg_CoastToLS` = spreadsheet spelling; `.Cfg_CoastTolS` =
      definitions spelling.)
- [ ] Record the correct spelling here and tell Muse — whichever file is
      wrong gets a one-line fix in a follow-up commit. Do not edit either
      file yourself mid-test.

## 4. Rollback — how to revert a bad write

If a write put a wrong value on the PLC:

1. In the webapp, the audit log (`audit/audit-YYYY-MM-DD.jsonl`) records
   every write with the tag, old value, and new value. Find the entry.
2. Re-stage the old value: edit the cell back to the old value (the
   staged panel shows old → new, so the old value is right there).
3. Write again with the typed PLC-name confirmation and confirm the tag
   reports `verified`.
4. Cross-check in Studio 5000.
5. If the webapp cannot connect at all, the same rollback can be done with
   the CLI: put the old value in the spreadsheet and run
   `python pypax.py <ip/slot> write <PLC>_ConfigTags.xlsx` — the filename
   must contain the PLC name or the CLI refuses.

## 5. If a write reports `mismatch` (or `failed`)

- [ ] **Do not re-write blindly.** A mismatch means the read-back after
      the write disagreed with the intended value. The tag stays staged
      in the panel — nothing is lost.
- [ ] Check the tag in Studio 5000: what value does the PLC actually
      hold now — the old value, the new value, or something else?
- [ ] Check the audit JSONL entry for that write: what was sent, and what
      did the read-back return?
- [ ] Common causes: another source is writing the same tag (HMI, logic,
      another operator); the tag is read-only in that state; a REAL value
      outside the 1e-6 tolerance (note the exact values — the tolerance
      may need adjusting, not the tag).
- [ ] If the PLC holds the intended value and the read-back was a transient
      glitch, re-write the same staged value once and confirm `verified`.
- [ ] If the PLC holds the old value and the write genuinely did not take,
      investigate before retrying (permissions, tag state, another writer).
- [ ] If a write reports `failed`, the write itself did not go through —
      check connectivity and the PLC state first.

## 6. Done — record the outcome

- [ ] Read-only smoke test: pass / fail — notes: ______________________
- [ ] Single-tag write + write-back: pass / fail — notes: ______________
- [ ] P_ValveMO spelling: ______________________
- [ ] Any mismatches/failures and what caused them: ____________________

Report the outcome back to Muse so the results (and the P_ValveMO fix)
can be recorded. The webapp is only "proven" once this checklist passes
on real hardware.
