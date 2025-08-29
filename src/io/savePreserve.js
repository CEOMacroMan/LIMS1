import { state } from '../state.js';
import { applyEdits } from '../features/applyEdits.js';
import { buildPreserveBinary, mimeType } from './zipPreserve.js';
import { logKV, log, DEBUG_MODE } from '../logger.js';
import { setStatus } from '../utils/dom.js';

export async function savePreserve() {
  if (!state.fsHandle) {
    setStatus('No file handle. Use Open for edit first.');
    return;
  }
  logKV('[save-preserve] selection', state.currentSelection);
  try {
    applyEdits();
    const ab = await buildPreserveBinary();
    const perm = await state.fsHandle.queryPermission({ mode: 'readwrite' });
    if (perm === 'prompt') await state.fsHandle.requestPermission({ mode: 'readwrite' });
    const w = await state.fsHandle.createWritable();
    await w.truncate(0);
    await w.write(ab);
    await w.close();
    log('[save-preserve] wrote file');
    setStatus('Saved');
  } catch (err) {
    logKV('[error]', { action: 'save-preserve', step: err.step || 'write', name: err.name, message: err.message });
    setStatus('Save failed');
  }
}

export async function downloadPreserve() {
  logKV('[download-preserve] selection', state.currentSelection);
  try {
    applyEdits();
    const ab = await buildPreserveBinary();
    const blob = new Blob([ab], { type: mimeType() });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = state.originalName || ('workbook.' + state.originalExt);
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setStatus('Download started');
  } catch (err) {
    logKV('[error]', { action: 'download-preserve', step: err.step || 'build', name: err.name, message: err.message });
    setStatus('Download failed');
  }
}
