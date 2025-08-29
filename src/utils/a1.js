/* A1 helpers using SheetJS utils */
export function encodeCell(rc) {
  return XLSX.utils.encode_cell(rc);
}

export function decodeRange(r) {
  return XLSX.utils.decode_range(r);
}

export function encodeRange(r) {
  return XLSX.utils.encode_range(r);
}
