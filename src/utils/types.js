import { log } from '../logger.js';

export function isNumericLiteral(str) {
  if (typeof str !== 'string') return false;
  if (str.trim() === '') return false;
  return !isNaN(Number(str));
}

export function normalizeForWrite(raw) {
  if (raw == null) return { kind: 'blank' };
  const t = typeof raw;
  if (t === 'number') {
    return Number.isNaN(raw) ? { kind: 'blank' } : { kind: 'number', v: raw };
  }
  if (t === 'boolean') return { kind: 'boolean', v: raw };
  if (t === 'string') {
    const trimmed = raw.trim();
    if (trimmed === '') return { kind: 'blank' };
    if (isNumericLiteral(trimmed)) return { kind: 'number', v: Number(trimmed) };
    return { kind: 'string', v: raw };
  }
  if (t === 'object') {
    let v;
    if (raw && typeof raw.v !== 'undefined') v = raw.v;
    else v = String(raw);
    return normalizeForWrite(v);
  }
  log('normalizeForWrite: treating value as string: ' + String(raw));
  return { kind: 'string', v: String(raw) };
}
