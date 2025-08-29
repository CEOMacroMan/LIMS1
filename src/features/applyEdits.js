/* global XLSX */
import { state } from '../state.js';
import { normalizeForWrite } from '../utils/types.js';
import { log } from '../logger.js';

export function applyEdits() {
  const { workbook, selection, editableData } = state;
  if (!workbook || !selection) return;
  const ws = workbook.Sheets[selection.sheet];
  let processed = 0;
  for (let r = 0; r < editableData.length; ++r) {
    const row = editableData[r];
    for (let c = 0; c < row.length; ++c) {
      const addr = XLSX.utils.encode_cell({ r: selection.start.r + r, c: selection.start.c + c });
      const norm = normalizeForWrite(row[c]);
      if (norm === null) {
        if (ws[addr]) delete ws[addr];
      } else {
        ws[addr] = norm;
      }
      processed++;
      if (processed % 100 === 0) log(`[applyEdits] processed ${processed} cells`);
    }
  }
}
