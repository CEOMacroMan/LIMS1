/* global fetch */
import { state } from '../state.js';
import { readWorkbook } from '../excel/workbook.js';
import { discoverTables } from '../excel/tables.js';
import { log, logKV } from '../logger.js';

export const fsSupported = ('showOpenFilePicker' in window) && (location.protocol === 'https:' || location.hostname === 'localhost');

async function loadFromArrayBuffer(ab, fname, handle) {
  state.workbook = readWorkbook(ab);
  state.originalName = fname;
  const parts = fname.split('.');
  state.originalExt = parts[parts.length - 1].toLowerCase();
  state.originalAb = ab;
  state.fileHandle = handle || null;
  state.editableData = [];
  state.selection = null;
  const tables = await discoverTables(ab, state.workbook);
  state.tableEntries = tables;
  logKV('[load] sheets', Object.keys(state.workbook.Sheets));
  return { tables };
}

export async function loadFromUrl(url) {
  const res = await fetch(url);
  const ab = await res.arrayBuffer();
  const name = url.split('/').pop() || 'workbook.xlsx';
  return await loadFromArrayBuffer(ab, name, null);
}

export async function openForEdit() {
  if (!fsSupported) return;
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false });
    const file = await handle.getFile();
    const ab = await file.arrayBuffer();
    return await loadFromArrayBuffer(ab, file.name, handle);
  } catch (err) {
    logKV('[error]', { action: 'open-for-edit', step: 'open', name: err.name, message: err.message });
    throw err;
  }
}
