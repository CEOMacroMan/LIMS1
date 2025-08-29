import { state } from '../state.js';

export function readWorkbook(ab) {
  state.sheetjsWb = XLSX.read(ab, { type: 'array' });
  return state.sheetjsWb;
}

export function writeWorkbookArrayBuffer() {
  const ext = state.originalExt;
  const opts = { bookType: ext, type: 'array' };
  if (ext === 'xlsm' || ext === 'xlsb') opts.bookVBA = true;
  return XLSX.write(state.sheetjsWb, opts);
}
