/* global XLSX */
export const A1 = {
  decodeRange: r => XLSX.utils.decode_range(r),
  encodeRange: r => XLSX.utils.encode_range(r),
  encodeCell: c => XLSX.utils.encode_cell(c)
};
