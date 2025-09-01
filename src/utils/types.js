export function isNumericLiteral(v) {
  return (typeof v === 'string') && v.trim() !== '' && !isNaN(v);
}

export function normalizeForWrite(v) {
  if (v === '') return null;
  if (isNumericLiteral(v)) return { t: 'n', v: Number(v) };
  if (typeof v === 'string' && /^true$/i.test(v)) return { t: 'b', v: true };
  if (typeof v === 'string' && /^false$/i.test(v)) return { t: 'b', v: false };
  return { t: 's', v: String(v) };
}
