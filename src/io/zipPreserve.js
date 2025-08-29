import { writeWorkbookArrayBuffer } from '../excel/workbook.js';
import { state } from '../state.js';
import { log, logKV } from '../logger.js';
import { MIME_MAP } from '../constants.js';

export async function buildPreserveBinary() {
  const ab = writeWorkbookArrayBuffer();
  logKV('[out-binary]', { ext: state.originalExt, byteLength: ab.byteLength });
  try {
    if (['xlsx', 'xlsm', 'xlsb', 'ods', 'numbers'].includes(state.originalExt)) {
      await JSZip.loadAsync(ab);
      log('[zip] ok for ' + state.originalExt);
    }
  } catch (e) {
    const err = new Error('zip validation failed');
    err.step = 'zip-validate';
    err.cause = e;
    throw err;
  }
  return ab;
}

export function mimeType() {
  return MIME_MAP[state.originalExt] || MIME_MAP.xlsx;
}
