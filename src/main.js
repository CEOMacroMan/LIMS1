import { state } from './state.js';
import { HEADER_TEXT } from './constants.js';
import { log, logKV } from './logger.js';
import { loadFromUrl, loadFromFile, openForEdit } from './io/loaders.js';
import { renderGrid } from './render/grid.js';
import { savePreserve, downloadPreserve } from './io/savePreserve.js';
import { downloadDataOnly } from './io/saveDataOnly.js';
import { setStatus, clear, el } from './utils/dom.js';
import { decodeRange, encodeRange } from './utils/a1.js';

function updateHeader() {
  document.querySelector('.app-header').textContent = HEADER_TEXT;
}

function updateSaveButton() {
  const fmtBtn = document.getElementById('saveFmtBtn');
  if (fmtBtn) fmtBtn.disabled = !state.fsHandle;
}

function populateTableList() {
  const sel = document.getElementById('tableList');
  clear(sel);
  state.tableEntries.forEach((t, i) => {
    const opt = el('option', { value: String(i), textContent: t.type === 'sheet' ? t.sheet + ' (preview)' : `${t.sheet}: ${t.name} [${t.ref}]` });
    sel.appendChild(opt);
  });
}

function showC1(sheetName) {
  const cellDiv = document.getElementById('cellC1');
  if (!state.sheetjsWb) { cellDiv.textContent = 'C1: (no workbook)'; return; }
  const ws = state.sheetjsWb.Sheets[sheetName || state.sheetjsWb.SheetNames[0]];
  if (!ws) { cellDiv.textContent = 'C1: (sheet missing)'; return; }
  const cell = ws['C1'];
  cellDiv.textContent = 'C1: ' + (cell ? cell.v : '(blank)');
}

function renderSelected() {
  const sel = document.getElementById('tableList');
  const idx = sel.selectedIndex;
  const info = state.tableEntries[idx];
  if (!info) { setStatus('Select a table'); return; }
  const ws = state.sheetjsWb.Sheets[info.sheet];
  if (!ws) { setStatus('Sheet not found'); return; }
  let rangeStr = info.ref;
  if (info.type === 'sheet') {
    const used = decodeRange(ws['!ref'] || 'A1');
    const r = { s: { r: used.s.r, c: used.s.c }, e: { r: Math.min(used.e.r, used.s.r + 49), c: Math.min(used.e.c, used.s.c + 49) } };
    rangeStr = encodeRange(r);
  }
  const decoded = decodeRange(rangeStr);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: rangeStr, defval: '' });
  state.currentSelection = { sheet: info.sheet, range: rangeStr, start: decoded.s, end: decoded.e };
  renderGrid(rows);
  showC1(info.sheet);
  setStatus('Rendered');
  logKV('[render] size', { rows: rows.length, cols: rows.reduce((m,r)=>Math.max(m,r.length),0) });
}

async function handleLoadUrl() {
  const url = document.getElementById('urlInput').value;
  await loadFromUrl(url);
  populateTableList();
  showC1();
  updateSaveButton();
}

async function handleLoadFile() {
  const file = document.getElementById('fileInput').files[0];
  if (!file) return;
  await loadFromFile(file);
  populateTableList();
  showC1();
  updateSaveButton();
}

async function handleOpenFs() {
  await openForEdit();
  populateTableList();
  showC1();
  updateSaveButton();
}

updateHeader();

// wire events
document.getElementById('loadUrlBtn').addEventListener('click', handleLoadUrl);
document.getElementById('loadFileBtn').addEventListener('click', handleLoadFile);
document.getElementById('openFsBtn').addEventListener('click', handleOpenFs);
document.getElementById('saveFmtBtn').addEventListener('click', savePreserve);
document.getElementById('downloadFmtBtn').addEventListener('click', downloadPreserve);
document.getElementById('downloadDataBtn').addEventListener('click', downloadDataOnly);
document.getElementById('renderBtn').addEventListener('click', renderSelected);

updateSaveButton();
