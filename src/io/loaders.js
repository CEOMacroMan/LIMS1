import { state } from '../state.js';
import { readWorkbook } from '../excel/workbook.js';
import { discoverStructures } from '../excel/tables.js';
import { log, logKV } from '../logger.js';
import { setStatus } from '../utils/dom.js';

function detectExt(name) {
  const lower = name.toLowerCase();
  if (lower.endsWith('.xlsm')) return 'xlsm';
  if (lower.endsWith('.xlsb')) return 'xlsb';
  if (lower.endsWith('.xls')) return 'biff8';
  if (lower.endsWith('.ods')) return 'ods';
  return 'xlsx';
}

async function loadArrayBuffer(ab, name) {
  readWorkbook(ab);
  state.tableEntries = await discoverStructures(ab);
  state.originalName = name;
  state.originalExt = detectExt(name);
  logKV('[load] sheets', state.sheetjsWb.SheetNames.length);
}

export async function loadFromUrl(url) {
  try {
    setStatus('Loading via URL...');
    const res = await fetch(url);
    const ab = await res.arrayBuffer();
    await loadArrayBuffer(ab, url.split('/').pop() || 'workbook.xlsx');
    setStatus('Loaded');
  } catch (err) {
    logKV('[error]', { action: 'load-url', message: err.message });
    setStatus('Failed to load URL');
  }
}

export async function loadFromFile(file) {
  try {
    setStatus('Loading file...');
    const ab = await file.arrayBuffer();
    await loadArrayBuffer(ab, file.name);
    setStatus('Loaded');
  } catch (err) {
    logKV('[error]', { action: 'load-file', message: err.message });
    setStatus('Failed to load file');
  }
}

export async function openForEdit() {
  if (!('showOpenFilePicker' in window)) {
    setStatus('FS Access not supported');
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false });
    const file = await handle.getFile();
    const ab = await file.arrayBuffer();
    state.fsHandle = handle;
    await loadArrayBuffer(ab, file.name);
    setStatus('File opened for edit');
  } catch (err) {
    logKV('[error]', { action: 'open-fs', message: err.message });
    setStatus('Open for edit failed');
  }
}
