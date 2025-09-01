import { state } from '../state.js';
import { buildPreserveBinary } from './zipPreserve.js';
import { MIME_MAP } from '../constants.js';
import { log, logKV } from '../logger.js';
import { setStatus } from '../utils/dom.js';
import { applyEdits } from '../features/applyEdits.js';

async function verifyPermission(handle) {
  const opts = { mode: 'readwrite' };
  if ((await handle.queryPermission(opts)) === 'granted') return true;
  if ((await handle.requestPermission(opts)) === 'granted') return true;
  return false;
}

export async function savePreserve() {
  const handle = state.fileHandle;
  if (!handle) return;
  logKV('[save]', { selection: state.selection });
  try {
    if (!(await verifyPermission(handle))) throw { step: 'permission', name: 'PermissionError', message: 'Permission denied' };
    applyEdits();
    const ab = await buildPreserveBinary();
    state.originalAb = ab;
    const w = await handle.createWritable();
    await w.truncate(0);
    await w.write(ab);
    await w.close();
    log('[save] done');
    setStatus('Saved');
  } catch (err) {
    logKV('[error]', { action: 'save', step: err.step || 'write', message: err.message });
    setStatus('Save failed');
    throw err;
  }
}

export async function downloadPreserve() {
  logKV('[download]', { selection: state.selection });
  try {
    applyEdits();
    const ab = await buildPreserveBinary();
    state.originalAb = ab;
    const mime = MIME_MAP[state.originalExt] || MIME_MAP.xlsx;
    const blob = new Blob([ab], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = state.originalName || `workbook.${state.originalExt}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    log('[download] initiated');
    setStatus('Download started');
  } catch (err) {
    logKV('[error]', { action: 'download', step: err.step || 'build-binary', message: err.message });
    setStatus('Download failed');
    throw err;
  }
}
