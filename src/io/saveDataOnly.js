/* global XLSX */
import { state } from '../state.js';
import { applyEdits } from '../features/applyEdits.js';
import { log } from '../logger.js';

export function downloadDataOnly() {
  applyEdits();
  const fname = state.originalName || `workbook.${state.originalExt}`;
  XLSX.writeFile(state.workbook, fname, { bookType: state.originalExt });
  log('Download (data-only) initiated');
}
