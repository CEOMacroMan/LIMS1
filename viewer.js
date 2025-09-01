/* global XLSX */
import { HEADER_TEXT } from './src/constants.js';
import { state } from './src/state.js';
import { log, logKV } from './src/logger.js';
import { setStatus } from './src/utils/dom.js';
import { loadFromUrl, openForEdit as openFs, fsSupported } from './src/io/loaders.js';
import { savePreserve, downloadPreserve } from './src/io/savePreserve.js';
import { downloadDataOnly } from './src/io/saveDataOnly.js';

import { populateTableSelect, enableSave } from './src/ui/controls.js';
import { renderGrid } from './src/render/grid.js';

// set header text
const header = document.querySelector('.app-header .title');
if (header) header.textContent = HEADER_TEXT;

function showC1(sheetName) {
  const cellDiv = document.getElementById('cellC1');
  if (!state.workbook) { cellDiv.textContent = 'C1: (no workbook)'; return; }
  const sheetNames = Object.keys(state.workbook.Sheets || {});
  let target = sheetName;
  if (!target) target = sheetNames.includes('INFO') ? 'INFO' : sheetNames[0];
  const ws = state.workbook.Sheets[target];
  const val = ws && ws['C1'] ? ws['C1'].v : '';
  cellDiv.textContent = String(val || '');
}

async function handleLoad(result) {
  state.tableEntries = result.tables;
  populateTableSelect(state.tableEntries);
  enableSave(Boolean(state.fileHandle));
  showC1();
  setStatus('Workbook loaded');
}

document.getElementById('loadUrlBtn').addEventListener('click', async () => {
  const url = document.getElementById('urlInput').value.trim();
  try {
    const res = await loadFromUrl(url);
    await handleLoad(res);
  } catch (err) {
    logKV('[error]', { action: 'load-url', step: err.step || 'fetch', name: err.name, message: err.message });
    setStatus('URL load failed');
  }
});

async function openForEdit() {
  log('[fs] openForEdit');
  if (!('showOpenFilePicker' in window)) {
    setStatus('File System Access API not supported in this browser.');
    document.getElementById('openFsBtn').disabled = true;
    document.getElementById('saveFmtBtn').disabled = true;
    return;
  }
  try {
    const res = await openFs();
    if (res) await handleLoad(res);
  } catch (err) {
    /* error already logged */
  }
}

document.getElementById('openFsBtn').addEventListener('click', openForEdit);

document.getElementById('saveFmtBtn').addEventListener('click', async () => {
  try { await savePreserve(); } catch (e) { /* logged */ }
});

document.getElementById('downloadBtn').addEventListener('click', async () => {
  try { await downloadPreserve(); } catch (e) { /* logged */ }
});

document.getElementById('downloadDataBtn').addEventListener('click', () => {
  downloadDataOnly();
});


document.getElementById('renderBtn').addEventListener('click', () => {
  renderSelected();
});

function renderSelected() {
  setStatus('');
  const sel = document.getElementById('tableList');
  const idx = sel.selectedIndex;
  const info = state.tableEntries[idx];
  if (!info || !state.workbook) {
    setStatus('Please load a workbook and select a table');
    return;
  }
  const ws = state.workbook.Sheets[info.sheet];
  if (!ws) { setStatus(`Sheet ${info.sheet} not found`); return; }
  let range;
  if (info.type === 'table' || info.type === 'name') {
    range = info.ref;
  } else {
    const used = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const end = { r: Math.min(used.e.r, used.s.r + 49), c: Math.min(used.e.c, used.s.c + 49) };
    range = XLSX.utils.encode_range({ s: used.s, e: end });
  }
  const dec = XLSX.utils.decode_range(range);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range, defval: '' });
  state.selection = { sheet: info.sheet, range, start: dec.s, end: dec.e };
  logKV('[render]', { sheet: info.sheet, name: info.name, ref: range });
  renderGrid(rows);
  const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
  log(`Rendered ${rows.length} rows and ${colCount} columns`);
  showC1(info.sheet);
}

if (!fsSupported) {
  document.getElementById('openFsBtn').disabled = true;
  document.getElementById('saveFmtBtn').disabled = true;
  setStatus('FS Access API requires HTTPS or localhost.');
}
