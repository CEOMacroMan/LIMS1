/* global XLSX */
export function readWorkbook(ab) {
  return XLSX.read(ab);
}

export function writeArrayBuffer(wb, ext) {
  const opts = { bookType: ext, type: 'array' };
  if (ext === 'xlsm' || ext === 'xlsb') opts.bookVBA = true;
  return XLSX.write(wb, opts);
}
