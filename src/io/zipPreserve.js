/* global JSZip */
import { state } from '../state.js';
import { writeArrayBuffer } from '../excel/workbook.js';
import { log, logKV } from '../logger.js';

const ZIP_EXTS = ['xlsx', 'xlsm', 'xlsb', 'ods', 'numbers'];

export async function buildPreserveBinary() {
  let ab;
  try {
    ab = writeArrayBuffer(state.workbook, state.originalExt);
    logKV('[out-binary]', { ext: state.originalExt, byteLength: ab.byteLength });
  } catch (err) {
    err.step = 'build-binary';
    throw err;
  }
  if (ZIP_EXTS.includes(state.originalExt)) {
    try {
      await JSZip.loadAsync(ab);
      log('[zip] ok for ' + state.originalExt);
    } catch (err) {
      err.step = 'zip-validate';
      throw err;
    }
  }
  return ab;
}
