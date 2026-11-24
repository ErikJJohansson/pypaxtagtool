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
 *   GET  /api/changes
 *   POST /api/write         {plc_name}
 *   GET  /api/audit
 *   GET  /api/export        (.xlsx download of baseline + staged)
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

/* Parse + validate a cell value against the spreadsheet data type
 * (row 6 of the AOI sheet). Returns {ok: true, value} or
 * {ok: false, reason}. */
function parseValue(raw, dataType) {
  const text = String(raw).trim();
  const type = String(dataType || '').toUpperCase();
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
  // STRING_* and anything else: keep as text.
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

const MockApi = {
  _connected: false,
  _staged: [],
  _audit: [],

  async health() {
    return { status: 'ok' };
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
    const lines = String(tsv).split(/\r?\n/);
    lines.forEach((line, lineIdx) => {
      if (line === '') return;
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
    return { applied, rejected, staged: this._staged.length };
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
    return { results, summary: { verified: results.length, mismatch: 0, failed: 0 } };
  },
  async getAudit() {
    return JSON.parse(JSON.stringify(this._audit));
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
  const profile = $('profile').value.trim() || 'default';
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
  updateBanner();
  showStatus('');
}

/* ------------------------------------------------------------------ */
/* AOI tabs + grid                                                     */
/* ------------------------------------------------------------------ */

async function loadAois() {
  state.aois = await api.listAois();
  const bar = $('tab-bar');
  bar.innerHTML = '';
  state.aois.forEach((aoi) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = aoi.name;
    const badge = document.createElement('span');
    badge.className = 'count-badge';
    badge.textContent = String(aoi.instance_count);
    btn.appendChild(badge);
    btn.addEventListener('click', () => openTab(aoi.name));
    bar.appendChild(btn);
  });
  if (state.aois.length > 0) {
    await openTab(state.aois[0].name);
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
    defs.push({
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
    });
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
    btn.classList.toggle('active', btn.textContent.startsWith(aoi));
  });
  state.activeAoi = aoi;
  if (!state.baselines[aoi]) {
    showStatus('Loading ' + aoi + '…');
    state.baselines[aoi] = await api.getAoi(aoi);
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
  $('modal-cancel-btn').addEventListener('click', () => {
    $('write-modal').hidden = true;
  });
  $('confirm-plc-input').addEventListener('input', () => {
    $('modal-confirm-btn').disabled =
      $('confirm-plc-input').value.trim() !== state.plcName;
  });
  $('modal-confirm-btn').addEventListener('click', confirmWrite);
  $('refresh-audit-btn').addEventListener('click', refreshAudit);
  updateBanner();

  if (!MOCK_MODE) {
    api.health().catch(() => {
      showStatus('Backend not reachable — start the backend, or open with ?mock=1 to try the UI with canned data.');
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
