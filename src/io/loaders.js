/* global fetch */
import { state } from '../state.js';
import { readWorkbook } from '../excel/workbook.js';
import { discoverTables } from '../excel/tables.js';
import { log, logKV } from '../logger.js';

export const fsSupported = ('showOpenFilePicker' in window) && (location.protocol === 'https:' || location.hostname === 'localhost');

async function loadFromArrayBuffer(ab, fname, handle) {
  try {
    state.workbook = readWorkbook(ab);
  } catch (e) {
    e.step = 'parse';
    throw e;
  }
  state.originalName = fname;
  const parts = fname.split('.');
  state.originalExt = parts[parts.length - 1].toLowerCase();
  state.originalAb = ab;
  state.fileHandle = handle || null;
  state.editableData = [];
  state.selection = null;
  log('[parse] sheets', state.workbook.SheetNames.join(', '));
  let discovery;
  try {
    discovery = await discoverTables(ab, state.workbook);
  } catch (e) {
    e.step = 'discover';
    throw e;
  }
  state.tableEntries = discovery.entries;
  state.exceljs = discovery.workbook;
  log('[tables]', discovery.tableCount, '[names]', discovery.nameCount);
  return { tables: discovery.entries };
}

export async function loadFromUrl(url) {
  const res = await fetch(url);
  const ab = await res.arrayBuffer();
  const name = url.split('/').pop() || 'workbook.xlsx';
  return await loadFromArrayBuffer(ab, name, null);
}

export async function openForEdit() {
  if (!fsSupported) return;
  let step = 'picker';
  try {
    log('[fs] picker');
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [{
        description: 'Excel',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
          'application/vnd.ms-excel.sheet.binary.macroEnabled.12': ['.xlsb'],
          'application/vnd.ms-excel': ['.xls']
        }
      }]
    });
    log('[fs] handle acquired');
    step = 'read';
    const file = await handle.getFile();
    logKV('[fs] file', { name: file.name, size: file.size });
    const ab = await file.arrayBuffer();
    step = 'parse';
    return await loadFromArrayBuffer(ab, file.name, handle);
  } catch (err) {
    logKV('[error]', { action: 'openForEdit', step: err.step || step, message: err.message });
    throw err;
  }
}
