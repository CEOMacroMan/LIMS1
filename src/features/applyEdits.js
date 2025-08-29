import { state } from '../state.js';
import { normalizeForWrite } from '../utils/types.js';
import { encodeCell } from '../utils/a1.js';
import { log, DEBUG_MODE } from '../logger.js';

function safeSample(v) {
  try {
    return typeof v === 'string' ? v.slice(0, 50) : JSON.stringify(v).slice(0, 120);
  } catch (e) {
    return String(v).slice(0, 120);
  }
}

export function applyEdits() {
  const { sheetjsWb, currentSelection, editableData } = state;
  if (!sheetjsWb || !currentSelection) return;
  const ws = sheetjsWb.Sheets[currentSelection.sheet];
  if (!ws) return;
  const errors = [];
  let processed = 0;
  for (let r = 0; r < editableData.length; ++r) {
    for (let c = 0; c < editableData[r].length; ++c) {
      const raw = editableData[r][c];
      const addr = encodeCell({ r: currentSelection.start.r + r, c: currentSelection.start.c + c });
      try {
        const norm = normalizeForWrite(raw);
        if (norm.kind === 'blank') {
          delete ws[addr];
        } else if (norm.kind === 'number') {
          ws[addr] = { t: 'n', v: norm.v };
        } else if (norm.kind === 'boolean') {
          ws[addr] = { t: 'b', v: norm.v };
        } else {
          ws[addr] = { t: 's', v: norm.v };
        }
      } catch (e) {
        errors.push({ r, c, addr, type: typeof raw, sample: safeSample(raw), message: e.message });
      }
      processed++;
      if (DEBUG_MODE && processed % 100 === 0) log(`[applyEdits] processed ${processed} cells`);
    }
  }
  if (errors.length) {
    log('[applyEdits] cell write errors: ' + JSON.stringify(errors.slice(0, 10)));
    const err = new Error('applyEdits failed for ' + errors.length + ' cells (see debug)');
    err.cells = errors;
    throw err;
  }
}
