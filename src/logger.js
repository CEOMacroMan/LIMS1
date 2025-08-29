export const DEBUG_MODE = true;

export function log(msg) {
  console.log(msg);
  const el = document.getElementById('debug');
  if (el) el.textContent += msg + '\n';
}

export function logKV(label, obj) {
  log(label + ': ' + (typeof obj === 'string' ? obj : JSON.stringify(obj)));
}
