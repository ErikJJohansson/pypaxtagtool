'use strict';
/* PyPaxTagTool web UI — vanilla JS, no build step, AG Grid Community 32.3.0.
 *
 * Talks to the frozen REST contract only (see PLAN.md / task brief):
 *   GET  /api/health
 *   POST /api/connect        {commpath, profile}
 *   POST /api/disconnect
 *   GET  /api/plc
 *   GET  /api/aois
 *   GET  /api/aois/{aoi}
 *   POST /api/aois/{aoi}/paste   {tsv, anchor_instance, anchor_subtag}
 *   POST /api/aois/import        multipart, form field `file` (.L5X)
 *   GET  /api/aois/scan
 *   POST /api/aois/import/confirm  {proposal_id, selected}
 *   DELETE /api/aois/custom/{name}
 *   GET  /api/changes
 *   POST /api/write         {plc_name}
 *   GET  /api/audit
 *   GET  /api/export        (.xlsx download of baseline + staged)
 *   POST /api/snapshots     {name?} — save the live baseline
 *   GET  /api/snapshots     — list, newest first
 *   GET  /api/snapshots/{name} — download the snapshot file
 *   POST /api/snapshots/upload — upload a snapshot file (.json)
 *   POST /api/snapshots/{name}/compare — diff vs the live PLC
 *   POST /api/snapshots/{name}/restore — stage differing values
 *   DELETE /api/snapshots/{name}
 *
 * CONTRACT CLARIFICATION (for the integrator to arbitrate with the backend):
 * the paste endpoint needs an anchor — the instance and subtag the pasted
 * block starts at — to map a pasted block onto the grid. A bare `tsv` cannot
 * express that. This UI therefore sends `anchor_instance` and
 * `anchor_subtag` alongside `tsv`. Single-cell edits are staged the same way,
 * as a one-cell TSV with an anchor. If the backend ever ignores the anchor,
 * the staged-changes panel (GET /api/changes, the source of truth for the
 * write modal) will disagree with the amber cells — that is the visible
 * symptom to watch for. The mock API in this file implements the anchor
 * semantics exactly, so mock mode is self-consistent.
 *
 * Mock mode: open with ?mock=1. All contract calls are answered from canned
 * in-memory data (MockApi below) — no server, no PLC.
 */

const MOCK_MODE = new URLSearchParams(window.location.search).get('mock') === '1';

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showStatus(message) {
  $('grid-status').textContent = message;
}

/* Longest string each type holds, in characters: STRING_N holds N, the
 * Logix STRING type holds 82, STRING_Area holds 8 (its UDT in the
 * P_ValveMO 4.10.06 export: DATA SINT[8]), STRING_Cond and STRING_NavTag
 * hold 20 each (UDTs in the P_Intlk 4.10.06 export: DATA SINT[20]).
 * Sizes learned from imported .L5X files arrive from
 * GET /api/string-limits at startup (learnedLimits below); until that
 * fetch lands, the pinned sizes here are the fallback. Returns null
 * when the size is unknown (that type is not length-checked). */
let learnedLimits = {};

function stringLimit(type) {
  if (type === 'STRING') return 82;
  if (type === 'STRING_AREA') return 8;
  if (type === 'STRING_COND' || type === 'STRING_NAVTAG') return 20;
  if (Object.prototype.hasOwnProperty.call(learnedLimits, type)) return learnedLimits[type];
  const match = /^STRING_(\d+)$/.exec(type);
  return match ? parseInt(match[1], 10) : null;
}

/* Parse + validate a cell value against the spreadsheet data type
 * (row 6 of the AOI sheet). Returns {ok: true, value} or
 * {ok: false, reason}. */
function parseValue(raw, dataType) {
  const text = String(raw).trim();
  const type = String(dataType || '').toUpperCase();
  if (!type) {
    // Untyped column: a custom AOI whose source .L5X is gone. Nothing is
    // accepted — the backend cannot validate it, and pycomm3 encodes any
    // truthy value as BOOL true, so a guess could write a wrong value.
    return { ok: false, reason: 'column type is unknown — re-import the .L5X to restore type checking' };
  }
  if (type.includes('BOOL')) {
    if (text === '1' || text.toLowerCase() === 'true') return { ok: true, value: 1 };
    if (text === '0' || text.toLowerCase() === 'false') return { ok: true, value: 0 };
    return { ok: false, reason: 'expected 0 or 1 for BOOL' };
  }
  if (type.includes('INT')) {
    if (/^-?\d+$/.test(text)) return { ok: true, value: parseInt(text, 10) };
    return { ok: false, reason: 'expected an integer for ' + dataType };
  }
  if (type.includes('REAL')) {
    if (text === '') return { ok: false, reason: 'expected a number for REAL' };
    const num = Number(text);
    if (Number.isFinite(num)) return { ok: true, value: num };
    return { ok: false, reason: 'expected a number for REAL' };
  }
  if (type.includes('STRING')) {
    // The spreadsheet flags a string longer than its type holds; the
    // PLC would not hold it either, so reject it like any other bad
    // value instead of staging a write that cannot land as-is.
    const limit = stringLimit(type);
    if (limit !== null && text.length > limit) {
      return { ok: false, reason: 'string is ' + text.length + ' characters, longer than the ' + limit + '-character limit for ' + dataType };
    }
    return { ok: true, value: text };
  }
  // Anything else: keep as text.
  return { ok: true, value: text };
}

/* ------------------------------------------------------------------ */
/* HTTP API client (live backend)                                      */
/* ------------------------------------------------------------------ */

const HttpApi = {
  async request(path, options) {
    const res = await fetch(path, {
      headers: { 'Content-Type': 'application/json' },
      ...options,
    });
    let body = null;
    try {
      body = await res.json();
    } catch (e) {
      body = null;
    }
    if (!res.ok) {
      const detail = (body && (body.detail || body.error)) || res.statusText;
      throw new Error(path + ' failed (' + res.status + '): ' + detail);
    }
    return body;
  },

  health() {
    return this.request('/api/health');
  },
  getStringLimits() {
    return this.request('/api/string-limits');
  },
  connect(commpath, profile) {
    return this.request('/api/connect', {
      method: 'POST',
      body: JSON.stringify({ commpath, profile }),
    });
  },
  disconnect() {
    return this.request('/api/disconnect', { method: 'POST', body: '{}' });
  },
  getPlc() {
    return this.request('/api/plc');
  },
  listAois() {
    return this.request('/api/aois');
  },
  getAoi(aoi) {
    return this.request('/api/aois/' + encodeURIComponent(aoi));
  },
  paste(aoi, tsv, anchorInstance, anchorSubtag) {
    return this.request('/api/aois/' + encodeURIComponent(aoi) + '/paste', {
      method: 'POST',
      body: JSON.stringify({ tsv, anchor_instance: anchorInstance, anchor_subtag: anchorSubtag }),
    });
  },
  getChanges() {
    return this.request('/api/changes');
  },
  discardChanges() {
    return this.request('/api/changes', { method: 'DELETE' });
  },
  write(plcName) {
    return this.request('/api/write', {
      method: 'POST',
      body: JSON.stringify({ plc_name: plcName }),
    });
  },
  getAudit() {
    return this.request('/api/audit');
  },
  /* Multipart upload — no JSON Content-Type; the browser sets the boundary. */
  async uploadImport(file) {
    const form = new FormData();
    form.append('file', file);
    const res = await fetch('/api/aois/import', { method: 'POST', body: form });
    let body = null;
    try {
      body = await res.json();
    } catch (e) {
      body = null;
    }
    if (!res.ok) {
      const detail = (body && (body.detail || body.error)) || res.statusText;
      throw new Error('Import upload failed (' + res.status + '): ' + detail);
    }
    return body;
  },
  scanAois() {
    return this.request('/api/aois/scan');
  },
  confirmImport(proposalId, selected) {
    return this.request('/api/aois/import/confirm', {
      method: 'POST',
      body: JSON.stringify({ proposal_id: proposalId, selected }),
    });
  },
  removeCustomAoi(name) {
    return this.request('/api/aois/custom/' + encodeURIComponent(name), {
      method: 'DELETE',
    });
  },
  listSnapshots() {
    return this.request('/api/snapshots');
  },
  saveSnapshot(name) {
    return this.request('/api/snapshots', {
      method: 'POST',
      body: JSON.stringify({ name: name || null }),
    });
  },
  /* Snapshot download — a file, not JSON: fetch the blob and trigger
   * a browser download (same pattern as exportSpreadsheet). */
  async downloadSnapshot(name) {
    const res = await fetch('/api/snapshots/' + encodeURIComponent(name));
    if (!res.ok) {
      let detail = res.statusText;
      try {
        const body = await res.json();
        if (body && body.detail) detail = body.detail;
      } catch (e) {
        /* keep the status text */
      }
      throw new Error('Snapshot download failed (' + res.status + '): ' + detail);
    }
    const blob = await res.blob();
    const filename = filenameFromDisposition(res.headers.get('Content-Disposition')) ||
      name + '.json';
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    return { name };
  },
  compareSnapshot(name) {
    return this.request('/api/snapshots/' + encodeURIComponent(name) + '/compare', {
      method: 'POST',
      body: '{}',
    });
  },
  restoreSnapshot(name) {
    return this.request('/api/snapshots/' + encodeURIComponent(name) + '/restore', {
      method: 'POST',
      body: '{}',
    });
  },
  deleteSnapshot(name) {
    return this.request('/api/snapshots/' + encodeURIComponent(name), {
      method: 'DELETE',
    });
  },
};

/* ------------------------------------------------------------------ */
/* Mock API — same shapes, canned in-memory data. Used with ?mock=1.    */
/* ------------------------------------------------------------------ */

const MOCK_DATA = {
  plcName: 'MOCK_PLC',
  commpath: '10.10.17.10/4',
  aois: [
    { name: 'P_AIn', instance_count: 3, subtag_count: 4 },
    { name: 'P_DIn', instance_count: 2, subtag_count: 3 },
  ],
  // plcData is the "PLC baseline": GET /api/aois/{aoi} returns this, and the
  // frontend overlays staged changes on top (see applyStagedOverlay).
  plcData: {
    P_AIn: {
      name: 'P_AIn',
      columns: [
        { subtag: '.Cfg_Desc', data_type: 'STRING_40', description: 'Description for display on HMI' },
        { subtag: '.Cfg_HiLim', data_type: 'REAL', description: 'High alarm limit' },
        { subtag: '.Cfg_LoLim', data_type: 'REAL', description: 'Low alarm limit' },
        { subtag: '.Cfg_HiHiLim', data_type: 'REAL', description: 'High-high alarm limit' },
      ],
      rows: [
        { name: 'FIT-101', values: { '.Cfg_Desc': 'Flow transmitter', '.Cfg_HiLim': 90, '.Cfg_LoLim': 10, '.Cfg_HiHiLim': 95 } },
        { name: 'PIT-201', values: { '.Cfg_Desc': 'Pressure transmitter', '.Cfg_HiLim': 150, '.Cfg_LoLim': 5, '.Cfg_HiHiLim': 160 } },
        { name: 'TIT-301', values: { '.Cfg_Desc': 'Temperature transmitter', '.Cfg_HiLim': 120, '.Cfg_LoLim': 0, '.Cfg_HiHiLim': 130 } },
      ],
    },
    P_DIn: {
      name: 'P_DIn',
      columns: [
        { subtag: '.Cfg_Desc', data_type: 'STRING_40', description: 'Description for display on HMI' },
        { subtag: '.Cfg_0StText', data_type: 'STRING_20', description: 'State 0 text' },
        { subtag: '.Cfg_1StText', data_type: 'STRING_20', description: 'State 1 text' },
      ],
      rows: [
        { name: 'XIC-301', values: { '.Cfg_Desc': 'Valve closed switch', '.Cfg_0StText': 'Closed', '.Cfg_1StText': 'Open' } },
        { name: 'XIC-302', values: { '.Cfg_Desc': 'Valve open switch', '.Cfg_0StText': 'Closed', '.Cfg_1StText': 'Open' } },
      ],
    },
  },
};

/* Canned import proposal for mock mode: a small custom valve AOI with one
 * nested-object group (FullStall), mirroring the frozen proposal shape.
 * The dev mock server (scripts/dev_mock_api.py) carries the same proposal. */
const MOCK_PROPOSAL = {
  aoi: 'P_MyValve',
  revision: '1.0.0',
  source_file: 'P_MyValve.L5X',
  columns: [
    { subtag: '.Cfg_Desc', type: 'STRING_40', description: 'Description for display on HMI', default: 'My Valve', selected: true },
    { subtag: '.Cfg_Label', type: 'STRING_20', description: 'Label for graphic symbol', default: 'My Valve', selected: true },
    { subtag: '.Cfg_HasOpenLS', type: 'BOOL', description: '1=Has open limit switch', default: 0, selected: true },
    { subtag: '.Cfg_OpenT', type: 'REAL', description: 'Time to open (sec)', default: 5, selected: true },
    { subtag: '.FullStall.Cfg_Exists', type: 'BOOL', description: '1=Full Stall Alarm exists', default: 0, selected: true },
    { subtag: '.FullStall.Cfg_Severity', type: 'DINT', description: 'Alarm Severity', default: 750, selected: true },
  ],
  nested_objects: ['FullStall'],
  proposal_id: 'mock-proposal-1',
};

/* A changed export of the canned valve, for mock-mode re-import: one
 * description changed (.Cfg_OpenT), one column added (.Cfg_NewCol),
 * one removed (.Cfg_HasOpenLS). The dev mock server carries the same. */
const MOCK_PROPOSAL_V2 = {
  aoi: 'P_MyValve',
  revision: '1.1.0',
  source_file: 'P_MyValve.L5X',
  columns: [
    { subtag: '.Cfg_Desc', type: 'STRING_40', description: 'Description for display on HMI', default: 'My Valve', selected: true },
    { subtag: '.Cfg_Label', type: 'STRING_20', description: 'Label for graphic symbol', default: 'My Valve', selected: true },
    { subtag: '.Cfg_OpenT', type: 'REAL', description: 'Time to fully open (sec)', default: 5, selected: true },
    { subtag: '.Cfg_NewCol', type: 'REAL', description: 'A newly added config value', default: 1.5, selected: true },
    { subtag: '.FullStall.Cfg_Exists', type: 'BOOL', description: '1=Full Stall Alarm exists', default: 0, selected: true },
    { subtag: '.FullStall.Cfg_Severity', type: 'DINT', description: 'Alarm Severity', default: 750, selected: true },
  ],
  nested_objects: ['FullStall'],
  proposal_id: 'mock-proposal-2',
};

/* Diff confirmed mock columns against a new proposal's columns — the
 * same rules as the backend's diff_columns (type, description,
 * default). */
function mockUpdateDiff(beforeColumns, afterColumns) {
  const before = {};
  beforeColumns.forEach((c) => { before[c.subtag] = c; });
  const added = [];
  const removed = [];
  const changed = [];
  let unchanged = 0;
  const asCol = (c, subtag) => ({
    subtag,
    type: c.type || c.data_type || '',
    description: c.description || '',
    default: c.default === undefined ? null : c.default,
  });
  afterColumns.forEach((col) => {
    const old = before[col.subtag];
    if (!old) {
      added.push(asCol(col, col.subtag));
      return;
    }
    const beforeCol = asCol(old, col.subtag);
    const afterCol = asCol(col, col.subtag);
    if (beforeCol.type === afterCol.type && beforeCol.description === afterCol.description &&
        beforeCol.default === afterCol.default) {
      unchanged += 1;
    } else {
      changed.push({ subtag: col.subtag, before: beforeCol, after: afterCol });
    }
  });
  Object.keys(before).forEach((subtag) => {
    if (!afterColumns.some((c) => c.subtag === subtag)) {
      removed.push(asCol(before[subtag], subtag));
    }
  });
  // The mock always knows the before columns, so details are known.
  return { added, removed, changed, unchanged, before_details_known: true };
}

/* One-line diff summary: "1 added, 1 removed, 2 changed, 9 unchanged". */
function diffSummary(diff) {
  const parts = [];
  if (diff.added.length > 0) parts.push(diff.added.length + ' added');
  if (diff.removed.length > 0) parts.push(diff.removed.length + ' removed');
  if (diff.changed.length > 0) parts.push(diff.changed.length + ' changed');
  parts.push(diff.unchanged + ' unchanged');
  return parts.join(', ');
}

const MockApi = {
  _connected: false,
  _staged: [],
  _audit: [],
  _proposals: {},
  _snapshots: [], // full snapshot dicts, newest first by created_at
  _autoSeq: 0, // monotonic suffix so same-second auto-snapshots stay unique

  async health() {
    return { status: 'ok' };
  },
  async getStringLimits() {
    return {};
  },
  async connect(commpath, profile) {
    this._connected = true;
    return { connected: true, plc_name: MOCK_DATA.plcName, commpath: commpath || MOCK_DATA.commpath };
  },
  async disconnect() {
    this._connected = false;
    return { connected: false };
  },
  async getPlc() {
    return {
      connected: this._connected,
      plc_name: this._connected ? MOCK_DATA.plcName : null,
      commpath: this._connected ? MOCK_DATA.commpath : null,
    };
  },
  async listAois() {
    return MOCK_DATA.aois;
  },
  async getAoi(aoi) {
    const data = MOCK_DATA.plcData[aoi];
    if (!data) throw new Error('Unknown AOI: ' + aoi);
    return JSON.parse(JSON.stringify(data));
  },
  async paste(aoi, tsv, anchorInstance, anchorSubtag) {
    const data = MOCK_DATA.plcData[aoi];
    if (!data) throw new Error('Unknown AOI: ' + aoi);
    const columns = data.columns.map((c) => c.subtag);
    const rows = data.rows;
    const startRow = rows.findIndex((r) => r.name === anchorInstance);
    const startCol = columns.indexOf(anchorSubtag);
    const applied = [];
    const rejected = [];
    // Leading/trailing blank lines are not part of the block; a blank
    // line inside it is an empty row that still advances the row
    // position (matches the backend's parse_tsv).
    const lines = String(tsv).split(/\r?\n/);
    while (lines.length > 0 && lines[0].trim() === '') lines.shift();
    while (lines.length > 0 && lines[lines.length - 1].trim() === '') lines.pop();
    lines.forEach((line, lineIdx) => {
      if (line.trim() === '') return;
      line.split('\t').forEach((cell, cellIdx) => {
        const row = rows[startRow + lineIdx];
        const subtag = columns[startCol + cellIdx];
        if (!row || !subtag) {
          rejected.push({ row: anchorInstance, column: subtag || '?', reason: 'paste extends past the edge of the grid' });
          return;
        }
        const colDef = data.columns.find((c) => c.subtag === subtag);
        const parsed = parseValue(cell, colDef.data_type);
        if (!parsed.ok) {
          rejected.push({ row: row.name, column: subtag, reason: parsed.reason });
          return;
        }
        const oldValue = this._stagedValue(aoi, row.name, subtag, data);
        this._upsertStaged(aoi, row, subtag, oldValue, parsed.value);
        applied.push({ row: row.name, column: subtag });
      });
    });
    // Same shape as the real paste endpoint: the staged changes for
    // this AOI (a list), not a count.
    return {
      applied,
      rejected,
      staged: JSON.parse(JSON.stringify(this._staged.filter((s) => s.aoi === aoi))),
    };
  },
  _stagedValue(aoi, instance, subtag, data) {
    const existing = this._staged.find(
      (s) => s.aoi === aoi && s.instance === instance && s.subtag === subtag,
    );
    if (existing) return existing.new_value;
    const row = data.rows.find((r) => r.name === instance);
    return row ? row.values[subtag] : null;
  },
  _upsertStaged(aoi, row, subtag, oldValue, newValue) {
    const existing = this._staged.find(
      (s) => s.aoi === aoi && s.instance === row.name && s.subtag === subtag,
    );
    const entry = {
      aoi,
      instance: row.name,
      subtag,
      tag: row.name + subtag,
      old_value: oldValue,
      new_value: newValue,
    };
    if (existing) {
      existing.new_value = newValue;
    } else {
      this._staged.push(entry);
    }
  },
  async getChanges() {
    return JSON.parse(JSON.stringify(this._staged));
  },
  async discardChanges() {
    const discarded = this._staged.length;
    this._staged = [];
    return { discarded };
  },
  async write(plcName) {
    if (plcName !== MOCK_DATA.plcName) {
      return { results: [], summary: { verified: 0, mismatch: 0, failed: this._staged.length } };
    }
    // Auto-snapshot the pre-write baseline (like the real backend),
    // keeping the newest 10 auto-snapshots.
    const autoSnap = this._buildSnapshot(true);
    this._autoSeq += 1;
    autoSnap._name = autoSnap._name + '-' + this._autoSeq;
    this._snapshots.push(autoSnap);
    const autoForPlc = this._snapshots.filter((s) => s._auto && s.plc_name === MOCK_DATA.plcName);
    while (autoForPlc.length > 10) {
      const oldest = autoForPlc.shift();
      this._snapshots.splice(this._snapshots.indexOf(oldest), 1);
    }
    const results = this._staged.map((s) => {
      const data = MOCK_DATA.plcData[s.aoi];
      const row = data.rows.find((r) => r.name === s.instance);
      if (row) row.values[s.subtag] = s.new_value;
      return { tag: s.tag, status: 'verified' };
    });
    this._audit.push({
      timestamp: new Date().toISOString(),
      plc_name: MOCK_DATA.plcName,
      action: 'write',
      verified: results.length,
      mismatch: 0,
      failed: 0,
    });
    this._staged = [];
    return {
      results,
      summary: { verified: results.length, mismatch: 0, failed: 0 },
      auto_snapshot: { ok: true, name: autoSnap._name },
    };
  },
  async getAudit() {
    return JSON.parse(JSON.stringify(this._audit));
  },
  /* Import stubs: upload/scan propose the canned P_MyValve; confirm
   * adds it as a custom AOI (empty rows — no PLC instances yet). Once
   * it is imported, upload/scan propose the changed v2 export as an
   * update (status "update_available" + diff), and confirm rewrites
   * it in place — mirroring the backend's re-import flow. */
  _updateResponse(proposal) {
    const current = MOCK_DATA.plcData[proposal.aoi];
    const diff = mockUpdateDiff(current ? current.columns : [], proposal.columns);
    return {
      ...JSON.parse(JSON.stringify(proposal)),
      status: 'update_available',
      update_available: true,
      diff,
    };
  },
  async uploadImport(file) {
    const alreadyImported = MOCK_DATA.aois.some((a) => a.name === MOCK_PROPOSAL.aoi);
    const proposal = JSON.parse(JSON.stringify(alreadyImported ? MOCK_PROPOSAL_V2 : MOCK_PROPOSAL));
    this._proposals[proposal.proposal_id] = proposal;
    if (alreadyImported) return this._updateResponse(proposal);
    return proposal;
  },
  async scanAois() {
    const skipped = [
      { file: 'P_AIn.L5X', reason: "AOI 'P_AIn' already exists as a template AOI" },
    ];
    if (MOCK_DATA.aois.some((a) => a.name === MOCK_PROPOSAL.aoi)) {
      const proposal = JSON.parse(JSON.stringify(MOCK_PROPOSAL_V2));
      this._proposals[proposal.proposal_id] = proposal;
      return { proposals: [], skipped, updates: [this._updateResponse(proposal)] };
    }
    const proposal = JSON.parse(JSON.stringify(MOCK_PROPOSAL));
    this._proposals[proposal.proposal_id] = proposal;
    return {
      proposals: [proposal],
      skipped,
      updates: [],
    };
  },
  async confirmImport(proposalId, selected) {
    const proposal = this._proposals[proposalId];
    if (!proposal) throw new Error('Unknown proposal: ' + proposalId);
    const existing = MOCK_DATA.aois.find((a) => a.name === proposal.aoi);
    if (existing) {
      // Re-import: all-or-nothing update, selected is ignored.
      const columns = proposal.columns;
      existing.subtag_count = columns.length;
      MOCK_DATA.plcData[proposal.aoi] = {
        name: proposal.aoi,
        columns: columns.map((c) => ({
          subtag: c.subtag,
          data_type: c.type,
          description: c.description,
          default: c.default,
        })),
        rows: (MOCK_DATA.plcData[proposal.aoi] || { rows: [] }).rows,
      };
      delete this._proposals[proposalId];
      return { aoi: proposal.aoi, columns: columns.length, updated: true };
    }
    const columns = proposal.columns.filter((c) => selected.includes(c.subtag));
    MOCK_DATA.aois.push({
      name: proposal.aoi,
      instance_count: 0,
      subtag_count: columns.length,
      custom: true,
    });
    MOCK_DATA.plcData[proposal.aoi] = {
      name: proposal.aoi,
      columns: columns.map((c) => ({
        subtag: c.subtag,
        data_type: c.type,
        description: c.description,
        default: c.default,
      })),
      rows: [],
    };
    delete this._proposals[proposalId];
    return { aoi: proposal.aoi, columns: columns.length };
  },
  async removeCustomAoi(name) {
    const idx = MOCK_DATA.aois.findIndex((a) => a.name === name);
    if (idx === -1 || !MOCK_DATA.aois[idx].custom) {
      throw new Error('Unknown custom AOI: ' + name);
    }
    MOCK_DATA.aois.splice(idx, 1);
    delete MOCK_DATA.plcData[name];
    return { removed: name };
  },
  /* Snapshot stubs (Phase 9a): in-memory snapshots built from the
   * canned PLC data, same shapes as the real backend. */
  _snapshotMeta(snap) {
    const aois = snap.aois;
    return {
      name: snap._name,
      plc_name: snap.plc_name,
      created_at: snap.created_at,
      aoi_count: Object.keys(aois).length,
      instance_count: Object.values(aois).reduce((n, inst) => n + Object.keys(inst).length, 0),
      auto: !!snap._auto,
    };
  },
  _buildSnapshot(auto) {
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const stamp = now.getFullYear() + pad(now.getMonth() + 1) + pad(now.getDate()) +
      '-' + pad(now.getHours()) + pad(now.getMinutes()) + pad(now.getSeconds());
    const aois = {};
    Object.entries(MOCK_DATA.plcData).forEach(([aoi, data]) => {
      aois[aoi] = {};
      data.rows.forEach((row) => {
        aois[aoi][row.name] = JSON.parse(JSON.stringify(row.values));
      });
    });
    return {
      format: 'pypax-snapshot/1',
      plc_name: MOCK_DATA.plcName,
      created_at: now.toISOString(),
      aoi_version: '4.10.6',
      aois,
      _name: auto ? MOCK_DATA.plcName + '_' + stamp : null,
      _auto: !!auto,
    };
  },
  async listSnapshots() {
    return this._snapshots
      .map((s) => this._snapshotMeta(s))
      .sort((a, b) => (a.created_at < b.created_at ? 1 : a.created_at > b.created_at ? -1 : 0));
  },
  async saveSnapshot(name) {
    if (name && !/^[A-Za-z0-9_-]+$/.test(name)) {
      throw new Error('Snapshot name must be [A-Za-z0-9_-]+, got ' + name);
    }
    if (name && this._snapshots.some((s) => s._name === name)) {
      throw new Error("Snapshot '" + name + "' already exists (409)");
    }
    const snap = this._buildSnapshot(!name);
    snap._name = name || snap._name;
    this._snapshots.push(snap);
    return this._snapshotMeta(snap);
  },
  async downloadSnapshot(name) {
    const snap = this._snapshots.find((s) => s._name === name);
    if (!snap) throw new Error('Unknown snapshot: ' + name);
    const blob = new Blob([JSON.stringify(snap, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = name + '.json';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    return { name };
  },
  _diffAgainst(snap) {
    const differences = [];
    const aois = new Set([...Object.keys(snap.aois), ...Object.keys(MOCK_DATA.plcData)]);
    aois.forEach((aoi) => {
      const snapRows = snap.aois[aoi] || {};
      const data = MOCK_DATA.plcData[aoi];
      const curRows = {};
      (data ? data.rows : []).forEach((row) => { curRows[row.name] = row.values; });
      const instances = new Set([...Object.keys(snapRows), ...Object.keys(curRows)]);
      instances.forEach((instance) => {
        const snapVals = snapRows[instance] || {};
        const curVals = curRows[instance] || {};
        const subtags = new Set([...Object.keys(snapVals), ...Object.keys(curVals)]);
        subtags.forEach((subtag) => {
          const snapValue = snapVals[subtag] === undefined ? null : snapVals[subtag];
          const curValue = curVals[subtag] === undefined ? null : curVals[subtag];
          if (snapValue !== curValue) {
            differences.push({ aoi, instance, subtag, snapshot_value: snapValue, current_value: curValue });
          }
        });
      });
    });
    differences.sort((a, b) => (a.aoi + a.instance + a.subtag < b.aoi + b.instance + b.subtag ? -1 : 1));
    return differences;
  },
  async compareSnapshot(name) {
    const snap = this._snapshots.find((s) => s._name === name);
    if (!snap) throw new Error('Unknown snapshot: ' + name);
    return { plc_name: snap.plc_name, differences: this._diffAgainst(snap) };
  },
  async restoreSnapshot(name) {
    const snap = this._snapshots.find((s) => s._name === name);
    if (!snap) throw new Error('Unknown snapshot: ' + name);
    if (snap.plc_name !== MOCK_DATA.plcName) {
      throw new Error("Snapshot is for PLC '" + snap.plc_name + "', connected to '" + MOCK_DATA.plcName + "' (409)");
    }
    let staged = 0;
    const skipped = [];
    Object.entries(snap.aois).forEach(([aoi, instances]) => {
      const data = MOCK_DATA.plcData[aoi];
      Object.entries(instances).forEach(([instance, values]) => {
        Object.entries(values).forEach(([subtag, snapValue]) => {
          const entry = { aoi, instance, subtag };
          if (snapValue === null || snapValue === undefined) {
            skipped.push({ ...entry, reason: 'no value in snapshot' });
            return;
          }
          if (!data) {
            skipped.push({ ...entry, reason: "unknown AOI '" + aoi + "'" });
            return;
          }
          const colDef = data.columns.find((c) => c.subtag === subtag);
          if (!colDef) {
            skipped.push({ ...entry, reason: "unknown subtag '" + subtag + "' in " + aoi });
            return;
          }
          const row = data.rows.find((r) => r.name === instance);
          if (!row) {
            skipped.push({ ...entry, reason: "instance '" + instance + "' not found in PLC" });
            return;
          }
          const parsed = parseValue(String(snapValue), colDef.data_type);
          if (!parsed.ok) {
            skipped.push({ ...entry, reason: parsed.reason });
            return;
          }
          const current = row.values[subtag];
          if (parsed.value === current) return;
          this._upsertStaged(aoi, row, subtag, current, parsed.value);
          staged += 1;
        });
      });
    });
    return { staged, skipped };
  },
  async deleteSnapshot(name) {
    const idx = this._snapshots.findIndex((s) => s._name === name);
    if (idx === -1) throw new Error('Unknown snapshot: ' + name);
    this._snapshots.splice(idx, 1);
    return { removed: name };
  },
};

const api = MOCK_MODE ? MockApi : HttpApi;

/* ------------------------------------------------------------------ */
/* State                                                               */
/* ------------------------------------------------------------------ */

const state = {
  connected: false,
  plcName: '',
  commpath: '',
  aois: [],
  activeAoi: null,
  // aoi -> {columns, rows} as last fetched from the backend (PLC baseline).
  baselines: {},
  changedCells: new Set(), // "aoi|instance|subtag"
  rejectedCells: new Map(), // "aoi|instance|subtag" -> reason
  staged: [],
};

let gridApi = null;

/* Custom-AOI import UI state. */
let importBusy = false;
let currentProposal = null;
let removeTarget = null;

function cellKey(aoi, instance, subtag) {
  return aoi + '|' + instance + '|' + subtag;
}

/* ------------------------------------------------------------------ */
/* Connection + header                                                 */
/* ------------------------------------------------------------------ */

function updateBanner() {
  const banner = $('plc-banner');
  if (state.connected) {
    banner.classList.remove('disconnected');
    banner.classList.add('connected');
    $('conn-status').textContent = 'Connected';
    $('plc-name').textContent = state.plcName;
    $('plc-commpath').textContent = '(' + state.commpath + ')';
  } else {
    banner.classList.add('disconnected');
    banner.classList.remove('connected');
    $('conn-status').textContent = 'Disconnected';
    $('plc-name').textContent = '';
    $('plc-commpath').textContent = '';
  }
  $('disconnect-btn').disabled = !state.connected;
  $('write-btn').disabled = !state.connected || state.staged.length === 0;
}

async function connect() {
  const commpath = $('commpath').value.trim();
  // Empty means "no profile" — never send a default name: the backend
  // 404s on an unknown profile, and there is no 'default' profile
  // unless the user configured one.
  const profile = $('profile').value.trim() || null;
  $('connect-error').textContent = '';
  if (!commpath) {
    $('connect-error').textContent = 'Enter a PLC comm path, e.g. 10.10.17.10/4.';
    return;
  }
  $('connect-btn').disabled = true;
  try {
    const res = await api.connect(commpath, profile);
    if (!res.connected) {
      throw new Error(res.error || 'backend refused the connection');
    }
    state.connected = true;
    state.plcName = res.plc_name;
    state.commpath = res.commpath || commpath;
    updateBanner();
    await loadAois();
    $('not-connected-note').hidden = true;
    $('tab-bar').hidden = false;
    $('grid-wrap').hidden = false;
    await refreshStaged();
    await refreshAudit();
    await refreshSnapshots();
    showStatus('Connected to ' + state.plcName + '. Pick an AOI tab.');
  } catch (err) {
    $('connect-error').textContent = 'Connect failed: ' + err.message;
  } finally {
    $('connect-btn').disabled = false;
  }
}

async function disconnect() {
  try {
    await api.disconnect();
  } catch (err) {
    showStatus('Disconnect failed: ' + err.message);
  }
  state.connected = false;
  state.plcName = '';
  state.commpath = '';
  state.aois = [];
  state.activeAoi = null;
  state.baselines = {};
  state.changedCells.clear();
  state.rejectedCells.clear();
  state.staged = [];
  if (gridApi) {
    gridApi.destroy();
    gridApi = null;
  }
  $('tab-bar').innerHTML = '';
  $('tab-bar').hidden = true;
  $('grid-wrap').hidden = true;
  $('not-connected-note').hidden = false;
  $('staged-list').innerHTML = '<li class="empty">No staged changes.</li>';
  $('staged-count').textContent = '0';
  $('audit-list').innerHTML = '<li class="empty">No writes recorded yet.</li>';
  $('snapshot-list').innerHTML = '<li class="empty">No snapshots saved yet.</li>';
  $('snapshot-diff').hidden = true;
  snapshotDiffName = null;
  setSnapshotStatus('');
  updateBanner();
  showStatus('');
}

/* ------------------------------------------------------------------ */
/* AOI tabs + grid                                                     */
/* ------------------------------------------------------------------ */

async function loadAois(openName) {
  state.aois = await api.listAois();
  const bar = $('tab-bar');
  bar.innerHTML = '';
  state.aois.forEach((aoi) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.dataset.aoi = aoi.name;
    btn.textContent = aoi.name;
    const badge = document.createElement('span');
    badge.className = 'count-badge';
    badge.textContent = String(aoi.instance_count);
    btn.appendChild(badge);
    if (aoi.custom) {
      const customBadge = document.createElement('span');
      customBadge.className = 'custom-badge';
      customBadge.textContent = 'custom';
      customBadge.title = 'Imported AOI — not part of the PlantPAX template';
      btn.appendChild(customBadge);
    }
    btn.addEventListener('click', () => openTab(aoi.name));
    bar.appendChild(btn);
  });
  const target = openName || (state.aois.length > 0 ? state.aois[0].name : null);
  if (target) {
    await openTab(target);
  }
}

function baselineValue(aoi, instance, subtag) {
  const base = state.baselines[aoi];
  if (!base) return undefined;
  const row = base.rows.find((r) => r.name === instance);
  return row ? row.values[subtag] : undefined;
}

function isChanged(aoi, instance, subtag) {
  return state.changedCells.has(cellKey(aoi, instance, subtag));
}

function buildColumnDefs(aoi, columns) {
  const defs = [
    {
      field: 'name',
      headerName: 'Tag Name',
      pinned: 'left',
      lockPinned: true,
      editable: false,
      width: 180,
      cellClass: 'tag-name-cell',
      filter: true,
    },
  ];
  columns.forEach((col) => {
    const isBool = String(col.data_type || '').toUpperCase().includes('BOOL');
    const def = {
      field: col.subtag,
      headerName: col.subtag,
      headerTooltip: col.subtag + ' (' + col.data_type + ')' + (col.description ? ' — ' + col.description : ''),
      editable: true,
      dataType: col.data_type,
      sortable: true,
      resizable: true,
      filter: true,
      valueParser: (params) => {
        const parsed = parseValue(params.newValue, params.colDef.dataType);
        if (!parsed.ok) {
          showStatus('Rejected: ' + parsed.reason + ' — value kept as ' + JSON.stringify(params.oldValue) + '.');
          return params.oldValue;
        }
        return parsed.value;
      },
      cellClassRules: {
        'cell-changed': (params) => isChanged(aoi, params.data.name, params.colDef.field),
        'cell-rejected': (params) => state.rejectedCells.has(cellKey(aoi, params.data.name, params.colDef.field)),
      },
      tooltipValueGetter: (params) => {
        const reason = state.rejectedCells.get(cellKey(aoi, params.data.name, params.colDef.field));
        return reason ? 'Rejected: ' + reason : null;
      },
    };
    if (isBool) {
      // BOOLs read as checkboxes, not 0/1 text: the display cell shows
      // a ticked box for 1, and the editor is AG Grid's checkbox editor
      // (single click opens it). The valueParser above still runs, so
      // the committed boolean becomes 1/0 before staging.
      def.cellRenderer = (params) => {
        const checked = params.value === 1 || params.value === true;
        return '<input type="checkbox" tabindex="-1" ' + (checked ? 'checked' : '') + ' disabled>';
      };
      def.cellEditor = 'agCheckboxCellEditor';
      def.singleClickEdit = true;
    }
    defs.push(def);
  });
  return defs;
}

function rowDataFromBaseline(aoi) {
  const base = state.baselines[aoi];
  return base.rows.map((row) => ({ name: row.name, ...row.values }));
}

/* Overlay staged changes (server is source of truth) onto freshly
 * fetched baseline rows, and mark the amber cells. */
function applyStagedOverlay(aoi) {
  const rows = rowDataFromBaseline(aoi);
  state.staged
    .filter((s) => s.aoi === aoi)
    .forEach((s) => {
      const row = rows.find((r) => r.name === s.instance);
      if (row) {
        row[s.subtag] = s.new_value;
        state.changedCells.add(cellKey(aoi, s.instance, s.subtag));
      }
    });
  return rows;
}

async function openTab(aoi) {
  document.querySelectorAll('#tab-bar button').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.aoi === aoi);
  });
  state.activeAoi = aoi;
  const aoiEntry = state.aois.find((a) => a.name === aoi);
  // Only imported AOIs can be removed; template AOIs have no remove control.
  $('remove-aoi-btn').hidden = !(aoiEntry && aoiEntry.custom);
  if (!state.baselines[aoi]) {
    showStatus('Loading ' + aoi + '…');
    try {
      state.baselines[aoi] = await api.getAoi(aoi);
    } catch (err) {
      showStatus('Could not load ' + aoi + ': ' + err.message);
      return;
    }
  }
  const base = state.baselines[aoi];
  const rows = applyStagedOverlay(aoi);

  if (gridApi) {
    gridApi.destroy();
    gridApi = null;
  }
  const gridDiv = $('grid');
  gridDiv.innerHTML = '';
  $('grid-title').textContent = aoi + ' — ' + base.rows.length + ' instances, ' + base.columns.length + ' subtags';
  showStatus('');

  gridApi = agGrid.createGrid(gridDiv, {
    columnDefs: buildColumnDefs(aoi, base.columns),
    rowData: rows,
    getRowId: (params) => params.data.name,
    defaultColDef: { resizable: true, sortable: true, filter: true },
    enableRangeSelection: true,
    rowSelection: 'multiple',
    onCellValueChanged: (params) => onCellValueChanged(aoi, params),
    processDataFromClipboard: (params) => onPaste(aoi, params),
  });
}

function onCellValueChanged(aoi, params) {
  if (params.source !== 'edit' && params.source !== 'paste') return;
  const instance = params.data.name;
  const subtag = params.colDef.field;
  const baseline = baselineValue(aoi, instance, subtag);
  const key = cellKey(aoi, instance, subtag);
  if (params.newValue === baseline) {
    state.changedCells.delete(key);
  } else {
    state.changedCells.add(key);
  }
  state.rejectedCells.delete(key);
  if (gridApi) gridApi.refreshCells({ force: true });
  // Paste blocks are staged as one POST by onPaste; single edits stage here.
  if (params.source === 'edit') {
    stageSingleEdit(aoi, instance, subtag, params.newValue).catch((err) => {
      showStatus('Could not stage edit: ' + err.message);
    });
  }
}

async function stageSingleEdit(aoi, instance, subtag, value) {
  const res = await api.paste(aoi, String(value), instance, subtag);
  applyPasteResult(aoi, res);
  await refreshStaged();
}

/* Paste: AG Grid applies the block locally (we return the data unchanged);
 * we separately POST the block to the paste endpoint for validation and
 * staging, then revert + flag any cells the backend rejected. */
function onPaste(aoi, params) {
  const focused = gridApi ? gridApi.getFocusedCell() : null;
  let anchorInstance = null;
  let anchorSubtag = null;
  if (focused) {
    const rowNode = gridApi.getDisplayedRowAtIndex(focused.rowIndex);
    if (rowNode) anchorInstance = rowNode.data.name;
    const column = gridApi.getColumn(focused.column.getColId ? focused.column.getColId() : focused.column);
    if (column) anchorSubtag = column.getColId();
  }
  const tsv = params.data.map((row) => row.join('\t')).join('\n');
  submitPaste(aoi, tsv, anchorInstance, anchorSubtag).catch((err) => {
    showStatus('Paste could not be staged on the server: ' + err.message +
      ' — values are in the grid, but stage them again or edit cells individually.');
  });
  return params.data;
}

async function submitPaste(aoi, tsv, anchorInstance, anchorSubtag) {
  const res = await api.paste(aoi, tsv, anchorInstance, anchorSubtag);
  applyPasteResult(aoi, res);
  await refreshStaged();
}

function resolveRejectedRow(aoi, rowRef) {
  if (typeof rowRef === 'string' && baselineValue(aoi, rowRef, '__none__') !== undefined) return rowRef;
  if (typeof rowRef === 'string') {
    const base = state.baselines[aoi];
    if (base && base.rows.some((r) => r.name === rowRef)) return rowRef;
  }
  if (typeof rowRef === 'number') {
    const base = state.baselines[aoi];
    if (base && base.rows[rowRef]) return base.rows[rowRef].name;
  }
  return null;
}

function applyPasteResult(aoi, res) {
  const rejected = (res && res.rejected) || [];
  rejected.forEach((rej) => {
    const instance = resolveRejectedRow(aoi, rej.row);
    if (!instance) return;
    const key = cellKey(aoi, instance, rej.column);
    state.rejectedCells.set(key, rej.reason || 'rejected by server');
    state.changedCells.delete(key);
    if (gridApi && state.activeAoi === aoi) {
      const node = gridApi.getRowNode(instance);
      if (node) {
        // Revert to the PLC baseline value; onCellValueChanged ignores
        // source 'api', so this does not re-stage.
        node.setDataValue(rej.column, baselineValue(aoi, instance, rej.column));
      }
      gridApi.refreshCells({ force: true });
    }
  });
  if (rejected.length > 0) {
    showStatus(rejected.length + ' pasted cell(s) rejected: ' +
      rejected.slice(0, 3).map((r) => r.column + ' — ' + r.reason).join('; '));
  }
}

/* ------------------------------------------------------------------ */
/* Custom AOI import (Phase 8)                                         */
/*                                                                     */
/* Upload (.L5X) or folder scan produces a proposal; the proposal is   */
/* reviewed here with per-column checkboxes and only becomes a tab     */
/* after Confirm sends {proposal_id, selected}. Cancel discards it     */
/* with no server change. A 409 (name collision) is shown verbatim.    */
/* ------------------------------------------------------------------ */

function setImportBusy(busy) {
  importBusy = busy;
  $('import-btn').disabled = busy;
  $('scan-btn').disabled = busy;
  $('import-file').disabled = busy;
  if (currentProposal) updateImportSelectedCount();
}

function setImportStatus(message, isError) {
  const el = $('import-status');
  el.textContent = message || '';
  el.classList.toggle('error', !!isError);
}

async function onImportFilePicked(event) {
  const file = event.target.files && event.target.files[0];
  // Reset so picking the same file again still fires a change event.
  event.target.value = '';
  if (!file) return;
  if (!/\.l5x$/i.test(file.name)) {
    setImportStatus('Only .L5X files can be imported.', true);
    return;
  }
  if (file.size > 5 * 1024 * 1024) {
    setImportStatus('File is larger than the 5 MB import limit.', true);
    return;
  }
  setImportBusy(true);
  setImportStatus('Uploading ' + file.name + '…');
  try {
    const proposal = await api.uploadImport(file);
    setImportStatus('');
    openProposalModal(proposal);
  } catch (err) {
    setImportStatus('Import failed: ' + err.message, true);
  } finally {
    setImportBusy(false);
  }
}

async function onScanClicked() {
  setImportBusy(true);
  setImportStatus('Scanning import folder…');
  try {
    const res = await api.scanAois();
    setImportStatus('');
    renderScanResults(res);
  } catch (err) {
    setImportStatus('Scan failed: ' + err.message, true);
  } finally {
    setImportBusy(false);
  }
}

function renderScanResults(res) {
  const section = $('scan-results');
  const proposalsEl = $('scan-proposals');
  const skippedEl = $('scan-skipped');
  proposalsEl.innerHTML = '';
  skippedEl.innerHTML = '';
  const proposals = (res && res.proposals) || [];
  const skipped = (res && res.skipped) || [];
  const updates = (res && res.updates) || [];
  if (proposals.length === 0 && skipped.length === 0 && updates.length === 0) {
    const p = document.createElement('p');
    p.className = 'empty';
    p.textContent = 'No .L5X files found in the import folder.';
    proposalsEl.appendChild(p);
  }
  proposals.forEach((proposal) => {
    const div = document.createElement('div');
    div.className = 'scan-proposal';
    const text = document.createElement('span');
    text.textContent = proposal.aoi + ' — rev ' + proposal.revision + ', ' +
      proposal.columns.length + ' columns, from ' + proposal.source_file;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = 'Review';
    btn.addEventListener('click', () => openProposalModal(proposal));
    div.appendChild(text);
    div.appendChild(btn);
    proposalsEl.appendChild(div);
  });
  updates.forEach((proposal) => {
    const div = document.createElement('div');
    div.className = 'scan-update';
    const badge = document.createElement('span');
    badge.className = 'update-badge';
    badge.textContent = 'Update available';
    const text = document.createElement('span');
    text.textContent = proposal.aoi + ' — ' + diffSummary(proposal.diff) +
      ', from ' + proposal.source_file;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = 'Review update';
    btn.addEventListener('click', () => openProposalModal(proposal));
    div.append(badge, text, btn);
    proposalsEl.appendChild(div);
  });
  skipped.forEach((entry) => {
    const div = document.createElement('div');
    div.className = 'scan-skipped';
    const file = document.createElement('strong');
    file.textContent = entry.file;
    const reason = document.createElement('span');
    reason.className = 'skip-reason';
    reason.textContent = ' — skipped: ' + entry.reason;
    div.appendChild(file);
    div.appendChild(reason);
    skippedEl.appendChild(div);
  });
  section.hidden = false;
}

function openProposalModal(proposal) {
  // Work on a copy: unchecking columns must not mutate the caller's data.
  currentProposal = JSON.parse(JSON.stringify(proposal));
  const isUpdate = currentProposal.status === 'update_available';
  $('import-modal-title').textContent = isUpdate ? 'Review AOI update' : 'Review imported AOI';
  $('import-select-actions').hidden = isUpdate;
  $('import-first-col-label').textContent = isUpdate ? 'Change' : 'Include';
  $('import-hint').textContent = isUpdate
    ? 'All-or-nothing: confirming replaces the AOI definition with the new export. ' +
      'Nothing is saved until you confirm — Cancel discards the proposal.'
    : 'Only checked columns become part of the imported AOI. ' +
      'Nothing is saved until you confirm — Cancel discards the proposal.';
  $('import-confirm-btn').textContent = isUpdate ? 'Update AOI' : 'Import AOI';
  const meta = $('import-meta');
  meta.innerHTML = '';
  [
    ['AOI', currentProposal.aoi],
    ['Revision', currentProposal.revision || '—'],
    ['Source file', currentProposal.source_file || '—'],
  ].forEach(([label, value]) => {
    const dt = document.createElement('dt');
    dt.textContent = label;
    const dd = document.createElement('dd');
    dd.textContent = value;
    meta.appendChild(dt);
    meta.appendChild(dd);
  });
  if (isUpdate) {
    $('import-nested').textContent = diffSummary(currentProposal.diff) +
      ' — confirming replaces the AOI definition with the new export.';
    renderImportDiff();
  } else {
    const nested = currentProposal.nested_objects || [];
    $('import-nested').textContent = nested.length > 0
      ? 'Nested objects found: ' + nested.join(', ') + ' — their Cfg_ members become dotted columns.'
      : 'No nested objects found.';
    renderImportColumns();
  }
  $('import-modal-error').textContent = '';
  $('import-modal').hidden = false;
}

/* Re-import diff table: added green, removed red, changed old → new.
 * No checkboxes — an update is all-or-nothing with a visible diff. */
function renderImportDiff() {
  const tbody = $('import-columns');
  tbody.innerHTML = '';
  const diff = currentProposal.diff || { added: [], removed: [], changed: [], unchanged: 0 };
  if (diff.before_details_known === false) {
    // A legacy import carries no confirmed column details, so type /
    // description / default changes cannot be detected — only added
    // and removed columns are shown. Say so, instead of letting the
    // "unchanged" count imply those columns were compared.
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 5;
    td.textContent = 'Note: type, description and default changes cannot be detected for this AOI — it was imported before those details were stored. Only added and removed columns are shown.';
    tr.appendChild(td);
    tbody.appendChild(tr);
  }
  const show = (value) => (value === null || value === undefined ? '—' : String(value));
  const arrow = (before, after) => (before === after ? show(before) : show(before) + ' → ' + show(after));
  const addRow = (kind, label, subtag, type, description, def) => {
    const tr = document.createElement('tr');
    tr.className = 'diff-' + kind;
    const tdKind = document.createElement('td');
    tdKind.className = 'diff-kind';
    tdKind.textContent = label;
    const tdSubtag = document.createElement('td');
    tdSubtag.className = 'subtag';
    tdSubtag.textContent = subtag;
    const tdType = document.createElement('td');
    tdType.className = 'type';
    tdType.textContent = type;
    const tdDesc = document.createElement('td');
    tdDesc.textContent = description;
    const tdDefault = document.createElement('td');
    tdDefault.textContent = def;
    tr.append(tdKind, tdSubtag, tdType, tdDesc, tdDefault);
    tbody.appendChild(tr);
  };
  diff.added.forEach((col) => {
    addRow('added', 'Added', col.subtag, col.type, col.description || '', show(col.default));
  });
  diff.removed.forEach((col) => {
    addRow('removed', 'Removed', col.subtag, col.type, col.description || '', show(col.default));
  });
  diff.changed.forEach((entry) => {
    addRow(
      'changed',
      'Changed',
      entry.subtag,
      arrow(entry.before.type, entry.after.type),
      arrow(entry.before.description, entry.after.description),
      arrow(entry.before.default, entry.after.default),
    );
  });
  updateImportSelectedCount();
}

function renderImportColumns() {
  const tbody = $('import-columns');
  tbody.innerHTML = '';
  currentProposal.columns.forEach((col) => {
    const tr = document.createElement('tr');
    if (col.selected === false) tr.classList.add('unchecked');
    const tdCheck = document.createElement('td');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = col.selected !== false;
    checkbox.setAttribute('aria-label', 'Include ' + col.subtag);
    checkbox.addEventListener('change', () => {
      col.selected = checkbox.checked;
      tr.classList.toggle('unchecked', !checkbox.checked);
      updateImportSelectedCount();
    });
    tdCheck.appendChild(checkbox);
    const tdSubtag = document.createElement('td');
    tdSubtag.className = 'subtag';
    tdSubtag.textContent = col.subtag;
    const tdType = document.createElement('td');
    tdType.className = 'type';
    tdType.textContent = col.type;
    const tdDesc = document.createElement('td');
    tdDesc.textContent = col.description || '';
    const tdDefault = document.createElement('td');
    tdDefault.textContent =
      col.default === null || col.default === undefined ? '—' : String(col.default);
    tr.append(tdCheck, tdSubtag, tdType, tdDesc, tdDefault);
    tbody.appendChild(tr);
  });
  updateImportSelectedCount();
}

function updateImportSelectedCount() {
  if (currentProposal && currentProposal.status === 'update_available') {
    // No checkboxes on a re-import: the count line doubles as the diff
    // summary, and the Update button is gated on busy only.
    $('import-selected-count').textContent = diffSummary(currentProposal.diff);
    $('import-confirm-btn').disabled = importBusy;
    return;
  }
  const selected = currentProposal.columns.filter((c) => c.selected !== false).length;
  $('import-selected-count').textContent =
    selected + ' of ' + currentProposal.columns.length + ' columns selected';
  $('import-confirm-btn').disabled = selected === 0 || importBusy;
}

function setAllImportColumns(checked) {
  if (!currentProposal) return;
  currentProposal.columns.forEach((c) => {
    c.selected = checked;
  });
  renderImportColumns();
}

async function confirmImportProposal() {
  if (!currentProposal) return;
  const isUpdate = currentProposal.status === 'update_available';
  // An update is all-or-nothing: no column selection is sent.
  const selected = isUpdate
    ? null
    : currentProposal.columns
      .filter((c) => c.selected !== false)
      .map((c) => c.subtag);
  if (!isUpdate && selected.length === 0) return;
  setImportBusy(true);
  $('import-modal-error').textContent = '';
  try {
    const res = await api.confirmImport(currentProposal.proposal_id, selected);
    const name = (res && res.aoi) || currentProposal.aoi;
    $('import-modal').hidden = true;
    currentProposal = null;
    setImportStatus(
      (res && res.updated ? 'Updated AOI ' : 'Imported AOI ') + name +
        ' (' + (res ? res.columns : selected.length) + ' columns).',
    );
    delete state.baselines[name];
    await loadAois(name);
  } catch (err) {
    // 409 name collision lands here with the server's message, verbatim.
    $('import-modal-error').textContent = 'Import failed: ' + err.message;
  } finally {
    setImportBusy(false);
    if (currentProposal) updateImportSelectedCount();
  }
}

function openRemoveModal() {
  removeTarget = state.activeAoi;
  $('remove-aoi-name').textContent = removeTarget;
  $('remove-modal-error').textContent = '';
  $('remove-modal').hidden = false;
}

async function confirmRemove() {
  if (!removeTarget) return;
  $('remove-confirm-btn').disabled = true;
  $('remove-modal-error').textContent = '';
  try {
    await api.removeCustomAoi(removeTarget);
    $('remove-modal').hidden = true;
    const removed = removeTarget;
    removeTarget = null;
    delete state.baselines[removed];
    if (state.activeAoi === removed) {
      state.activeAoi = null;
      if (gridApi) {
        gridApi.destroy();
        gridApi = null;
      }
    }
    await loadAois();
    if (state.aois.length === 0) {
      $('tab-bar').hidden = true;
      $('grid-wrap').hidden = true;
    }
    setImportStatus('Removed custom AOI ' + removed + '.');
  } catch (err) {
    $('remove-modal-error').textContent = 'Remove failed: ' + err.message;
  } finally {
    $('remove-confirm-btn').disabled = false;
  }
}

/* ------------------------------------------------------------------ */
/* PLC snapshots (Phase 9a)                                            */
/*                                                                     */
/* Save the live baseline, list snapshots, compare one against the     */
/* live PLC (freshly re-read by the backend), or restore it — restore  */
/* only stages the differing values, nothing is written until the      */
/* user goes through the normal write confirmation.                    */
/* ------------------------------------------------------------------ */

let snapshotDiffName = null;

function setSnapshotStatus(message, isError) {
  const el = $('snapshot-status');
  el.textContent = message || '';
  el.classList.toggle('error', !!isError);
}

async function refreshSnapshots() {
  try {
    const entries = await api.listSnapshots();
    renderSnapshotList(entries);
  } catch (err) {
    setSnapshotStatus('Could not load snapshots: ' + err.message, true);
  }
}

function renderSnapshotList(entries) {
  const list = $('snapshot-list');
  list.innerHTML = '';
  if (!entries || entries.length === 0) {
    list.innerHTML = '<li class="empty">No snapshots saved yet.</li>';
  } else {
    entries.forEach((entry) => {
      const li = document.createElement('li');
      li.className = 'snapshot-item';
      const name = document.createElement('strong');
      name.className = 'snapshot-name';
      name.textContent = entry.name;
      const meta = document.createElement('span');
      meta.className = 'snapshot-meta';
      meta.textContent = (entry.auto ? 'auto · ' : '') + entry.plc_name + ' · ' +
        entry.created_at + ' · ' + entry.aoi_count + ' AOIs, ' +
        entry.instance_count + ' instances';
      const actions = document.createElement('span');
      actions.className = 'snapshot-actions';
      [
        ['Compare', () => onCompareSnapshot(entry.name)],
        ['Restore', () => onRestoreSnapshot(entry.name)],
        ['Download', () => onDownloadSnapshot(entry.name)],
        ['Delete', () => onDeleteSnapshot(entry.name)],
      ].forEach(([label, handler]) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = label;
        btn.addEventListener('click', handler);
        actions.appendChild(btn);
      });
      li.append(name, meta, actions);
      list.appendChild(li);
    });
  }
  $('snapshot-save-btn').disabled = !state.connected;
  if (snapshotDiffName && !(entries || []).some((e) => e.name === snapshotDiffName)) {
    $('snapshot-diff').hidden = true;
    snapshotDiffName = null;
  }
}

async function onSaveSnapshotClicked() {
  const raw = window.prompt(
    'Snapshot name (letters, digits, - and _ only; leave empty for an automatic name):',
  );
  if (raw === null) return; // cancelled
  const name = raw.trim();
  if (name !== '' && !/^[A-Za-z0-9_-]+$/.test(name)) {
    setSnapshotStatus('Snapshot name must be letters, digits, - or _ only.', true);
    return;
  }
  setSnapshotStatus('Saving snapshot…');
  try {
    const res = await api.saveSnapshot(name === '' ? null : name);
    setSnapshotStatus('Snapshot saved: ' + res.name + '.');
    await refreshSnapshots();
  } catch (err) {
    setSnapshotStatus('Save failed: ' + err.message, true);
  }
}

function formatSnapshotValue(value) {
  return value === null || value === undefined ? '— (missing)' : String(value);
}

function renderSnapshotDiff(name, differences) {
  snapshotDiffName = name;
  const box = $('snapshot-diff');
  box.innerHTML = '';
  const heading = document.createElement('p');
  if (differences.length === 0) {
    heading.textContent = 'Snapshot ' + name + ' matches the live PLC — no differences.';
    box.appendChild(heading);
  } else {
    heading.innerHTML = '<strong>' + escapeHtml(differences.length) +
      ' difference(s) vs the live PLC</strong> in snapshot ' + escapeHtml(name) + ':';
    box.appendChild(heading);
    const wrap = document.createElement('div');
    wrap.className = 'import-table-wrap';
    const table = document.createElement('table');
    table.className = 'import-table';
    table.innerHTML = '<thead><tr><th scope="col">AOI</th><th scope="col">Instance</th>' +
      '<th scope="col">Subtag</th><th scope="col">Snapshot</th><th scope="col">Live PLC</th></tr></thead>';
    const tbody = document.createElement('tbody');
    differences.forEach((d) => {
      const tr = document.createElement('tr');
      [d.aoi, d.instance, d.subtag].forEach((text) => {
        const td = document.createElement('td');
        td.className = 'subtag';
        td.textContent = text;
        tr.appendChild(td);
      });
      const tdSnap = document.createElement('td');
      tdSnap.className = 'old';
      tdSnap.textContent = formatSnapshotValue(d.snapshot_value);
      const tdCur = document.createElement('td');
      tdCur.className = 'new';
      tdCur.textContent = formatSnapshotValue(d.current_value);
      tr.append(tdSnap, tdCur);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    box.appendChild(wrap);
    const restoreBtn = document.createElement('button');
    restoreBtn.type = 'button';
    restoreBtn.textContent = 'Restore these values (stages them — nothing is written yet)';
    restoreBtn.addEventListener('click', () => onRestoreSnapshot(name));
    box.appendChild(restoreBtn);
  }
  box.hidden = false;
}

async function onCompareSnapshot(name) {
  setSnapshotStatus('Comparing ' + name + ' with the live PLC…');
  try {
    const res = await api.compareSnapshot(name);
    setSnapshotStatus('');
    renderSnapshotDiff(name, (res && res.differences) || []);
    if (res && res.plc_name && res.plc_name !== state.plcName) {
      // The snapshot belongs to another PLC: the diff is real, but
      // every value will differ. Say so up front.
      const box = $('snapshot-diff');
      const warn = document.createElement('p');
      warn.textContent = 'Note: this snapshot is for PLC ' + res.plc_name +
        ', but you are connected to ' + state.plcName + ' — every value will differ.';
      box.insertBefore(warn, box.firstChild);
    }
  } catch (err) {
    setSnapshotStatus('Compare failed: ' + err.message, true);
  }
}

async function onRestoreSnapshot(name) {
  setSnapshotStatus('Restoring ' + name + '…');
  try {
    const res = await api.restoreSnapshot(name);
    const skipped = (res && res.skipped) || [];
    let message = 'Restore staged ' + (res ? res.staged : 0) + ' change(s).';
    if (skipped.length > 0) {
      message += ' ' + skipped.length + ' skipped: ' +
        skipped.slice(0, 3).map((s) => s.subtag + ' — ' + s.reason).join('; ');
    }
    message += ' Nothing is written until you write.';
    setSnapshotStatus(message, skipped.length > 0);
    await refreshStaged();
    if (state.activeAoi) await openTab(state.activeAoi);
  } catch (err) {
    setSnapshotStatus('Restore failed: ' + err.message, true);
  }
}

async function onDownloadSnapshot(name) {
  setSnapshotStatus('Downloading ' + name + '…');
  try {
    await api.downloadSnapshot(name);
    setSnapshotStatus('Snapshot downloaded: ' + name + '.json');
  } catch (err) {
    setSnapshotStatus('Download failed: ' + err.message, true);
  }
}

async function onDeleteSnapshot(name) {
  if (!window.confirm('Delete snapshot ' + name + '? This cannot be undone.')) return;
  setSnapshotStatus('Deleting ' + name + '…');
  try {
    await api.deleteSnapshot(name);
    setSnapshotStatus('Snapshot deleted: ' + name + '.');
    await refreshSnapshots();
  } catch (err) {
    setSnapshotStatus('Delete failed: ' + err.message, true);
  }
}

/* ------------------------------------------------------------------ */
/* Staged changes + write + audit                                     */
/* ------------------------------------------------------------------ */

async function refreshStaged() {
  try {
    state.staged = await api.getChanges();
  } catch (err) {
    showStatus('Could not load staged changes: ' + err.message);
    return;
  }
  $('staged-count').textContent = String(state.staged.length);
  const list = $('staged-list');
  list.innerHTML = '';
  if (state.staged.length === 0) {
    list.innerHTML = '<li class="empty">No staged changes.</li>';
  } else {
    state.staged.forEach((s) => {
      const li = document.createElement('li');
      li.innerHTML = '<strong>' + escapeHtml(s.aoi) + '</strong> — ' +
        escapeHtml(s.instance) + escapeHtml(s.subtag) + ': ' +
        '<span class="old">' + escapeHtml(s.old_value) + '</span> → ' +
        '<span class="new">' + escapeHtml(s.new_value) + '</span>';
      list.appendChild(li);
    });
  }
  $('write-btn').disabled = !state.connected || state.staged.length === 0;
}

function openWriteModal() {
  const counts = {};
  state.staged.forEach((s) => {
    counts[s.aoi] = (counts[s.aoi] || 0) + 1;
  });
  $('modal-plc-name').textContent = state.plcName;
  $('modal-commpath').textContent = state.commpath;
  const ul = $('modal-counts');
  ul.innerHTML = '';
  Object.entries(counts).forEach(([aoi, count]) => {
    const li = document.createElement('li');
    li.textContent = aoi + ': ' + count + ' tag change' + (count === 1 ? '' : 's');
    ul.appendChild(li);
  });
  $('confirm-plc-input').value = '';
  $('modal-confirm-btn').disabled = true;
  $('modal-error').textContent = '';
  $('write-modal').hidden = false;
  $('confirm-plc-input').focus();
}

async function confirmWrite() {
  const typed = $('confirm-plc-input').value.trim();
  if (typed !== state.plcName) {
    $('modal-error').textContent = 'PLC name does not match — write cancelled.';
    return;
  }
  $('modal-confirm-btn').disabled = true;
  $('modal-error').textContent = '';
  try {
    const res = await api.write(state.plcName);
    $('write-modal').hidden = true;
    renderWriteResults(res);
    // Re-fetch baselines: written values are now the PLC values.
    state.baselines = {};
    state.changedCells.clear();
    state.rejectedCells.clear();
    await refreshStaged();
    await refreshAudit();
    if (state.activeAoi) await openTab(state.activeAoi);
  } catch (err) {
    $('modal-error').textContent = 'Write failed: ' + err.message;
    $('modal-confirm-btn').disabled = false;
  }
}

/* Handles both vocabularies: Phase 4 "written" and Phase 5
 * "verified" / "mismatch" / "failed". */
function renderWriteResults(res) {
  const box = $('write-results');
  const results = (res && res.results) || [];
  const summary = (res && res.summary) || {};
  const counts = { written: 0, verified: 0, mismatch: 0, failed: 0 };
  results.forEach((r) => {
    if (counts[r.status] !== undefined) counts[r.status] += 1;
  });
  const writtenTotal = (summary.written || 0) + (counts.verified || 0);
  let html = '<p><strong>Write finished.</strong> ' +
    writtenTotal + ' written/verified, ' +
    (counts.mismatch || 0) + ' mismatch, ' +
    (summary.failed || counts.failed || 0) + ' failed.</p>';
  // Every write first takes a pre-write auto-snapshot (Phase 9a) — say
  // which one, so the rollback point is visible. A failed capture is
  // reported, not hidden: the write still went ahead.
  if (res && res.auto_snapshot) {
    if (res.auto_snapshot.ok) {
      html += '<p>Pre-write snapshot: ' + escapeHtml(res.auto_snapshot.name) + '.</p>';
    } else {
      html += '<p>Pre-write snapshot failed: ' +
        escapeHtml(res.auto_snapshot.error || 'unknown error') +
        ' — the write still went ahead.</p>';
    }
  }
  if (results.length > 0) {
    html += '<ul>' + results.map((r) => {
      const cls = r.status === 'failed' ? 'status-failed'
        : r.status === 'mismatch' ? 'status-mismatch'
        : r.status === 'verified' ? 'status-verified' : 'status-written';
      return '<li><span class="status-chip ' + cls + '">' + escapeHtml(r.status) +
        '</span> ' + escapeHtml(r.tag) + '</li>';
    }).join('') + '</ul>';
  }
  box.innerHTML = html;
  box.hidden = false;
}

async function refreshAudit() {
  try {
    const entries = await api.getAudit();
    const list = $('audit-list');
    list.innerHTML = '';
    if (!entries || entries.length === 0) {
      list.innerHTML = '<li class="empty">No writes recorded yet.</li>';
      return;
    }
    entries.slice().reverse().forEach((e) => {
      const li = document.createElement('li');
      const when = e.timestamp ? escapeHtml(e.timestamp) : '';
      const plc = e.plc_name ? escapeHtml(e.plc_name) : '';
      const detail = e.action ? escapeHtml(e.action) : '';
      const written = e.written !== undefined ? ' — ' + e.written + ' written' : '';
      const failed = e.failed ? ', ' + e.failed + ' failed' : '';
      li.innerHTML = when + ' ' + plc + ' ' + detail + written + failed;
      list.appendChild(li);
    });
  } catch (err) {
    showStatus('Could not load audit log: ' + err.message);
  }
}

/* Download the filename from a Content-Disposition header, if any. */
function filenameFromDisposition(header) {
  if (!header) return null;
  const match = /filename="([^"]+)"/.exec(header);
  return match ? match[1] : null;
}

/* Export: fetch the spreadsheet from the backend and trigger a
 * browser download. The backend serves baseline values plus staged
 * changes (amber) as <plc_name>_ConfigTags.xlsx. */
async function exportSpreadsheet() {
  showStatus('Exporting spreadsheet…');
  let response;
  try {
    response = await fetch('/api/export');
  } catch (err) {
    showStatus('Export failed: backend not reachable.');
    return;
  }
  if (!response.ok) {
    let detail = response.statusText;
    try {
      const body = await response.json();
      if (body && body.detail) detail = body.detail;
    } catch (err) {
      /* keep the status text */
    }
    showStatus('Export failed: ' + detail);
    return;
  }
  const blob = await response.blob();
  const filename = filenameFromDisposition(
    response.headers.get('Content-Disposition'),
  ) || 'PyPax_ConfigTags.xlsx';
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  showStatus('Spreadsheet exported: ' + filename);
}

/* ------------------------------------------------------------------ */
/* Wiring + init                                                       */
/* ------------------------------------------------------------------ */

function init() {
  if (MOCK_MODE) {
    $('mock-badge').hidden = false;
  }
  $('connect-btn').addEventListener('click', connect);
  $('disconnect-btn').addEventListener('click', disconnect);
  $('refresh-staged-btn').addEventListener('click', refreshStaged);
  $('clear-highlight-btn').addEventListener('click', async () => {
    // Discard all staged changes on the server, then clear the local
    // amber/red highlighting and re-sync from the server.
    let discarded = 0;
    try {
      const res = await api.discardChanges();
      discarded = res.discarded || 0;
    } catch (err) {
      showStatus('Could not discard staged changes: ' + err.message);
      return;
    }
    state.changedCells.clear();
    state.rejectedCells.clear();
    if (gridApi) gridApi.refreshCells({ force: true });
    await refreshStaged();
    showStatus('Discarded ' + discarded + ' staged change(s).');
  });
  $('write-btn').addEventListener('click', openWriteModal);
  $('export-btn').addEventListener('click', exportSpreadsheet);
  $('import-btn').addEventListener('click', () => $('import-file').click());
  $('import-file').addEventListener('change', onImportFilePicked);
  $('scan-btn').addEventListener('click', onScanClicked);
  $('import-select-all-btn').addEventListener('click', () => setAllImportColumns(true));
  $('import-select-none-btn').addEventListener('click', () => setAllImportColumns(false));
  $('import-cancel-btn').addEventListener('click', () => {
    $('import-modal').hidden = true;
    currentProposal = null;
  });
  $('import-confirm-btn').addEventListener('click', confirmImportProposal);
  $('remove-aoi-btn').addEventListener('click', openRemoveModal);
  $('remove-cancel-btn').addEventListener('click', () => {
    $('remove-modal').hidden = true;
    removeTarget = null;
  });
  $('remove-confirm-btn').addEventListener('click', confirmRemove);
  $('modal-cancel-btn').addEventListener('click', () => {
    $('write-modal').hidden = true;
  });
  $('confirm-plc-input').addEventListener('input', () => {
    $('modal-confirm-btn').disabled =
      $('confirm-plc-input').value.trim() !== state.plcName;
  });
  $('modal-confirm-btn').addEventListener('click', confirmWrite);
  $('refresh-audit-btn').addEventListener('click', refreshAudit);
  $('snapshot-save-btn').addEventListener('click', onSaveSnapshotClicked);
  $('refresh-snapshots-btn').addEventListener('click', refreshSnapshots);
  updateBanner();

  // Learned string sizes: fetched once at startup. Until it lands the
  // pinned map in stringLimit is the fallback, so validation never
  // blocks on this fetch; a failed fetch just keeps the pinned sizes.
  api.getStringLimits().then((limits) => {
    learnedLimits = limits || {};
  }).catch(() => {});

  if (!MOCK_MODE) {
    api.health().catch(() => {
      showStatus('Backend not reachable — start the backend, or open with ?mock=1 to try the UI with canned data.');
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
