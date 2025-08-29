import { state } from '../state.js';
import { applyEdits } from '../features/applyEdits.js';
import { logKV } from '../logger.js';

export function downloadDataOnly() {
  if (!state.sheetjsWb) return;
  try {
    applyEdits();
    const opts = { bookType: state.originalExt, bookSST: true };
    XLSX.writeFile(state.sheetjsWb, state.originalName || 'workbook.' + state.originalExt, opts);
    logKV('[download-data]', { ext: state.originalExt });
  } catch (err) {
    logKV('[error]', { action: 'download-data', step: 'writeFile', name: err.name, message: err.message });
  }
}
