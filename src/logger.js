export const DEBUG_MODE = true;

export function log(msg) {
  if (!DEBUG_MODE) return;
  console.log(msg);
  const el = document.getElementById('debug');
  if (el) el.textContent += msg + '\n';
}

export function logKV(label, obj) {
  log(`${label} ${JSON.stringify(obj)}`);
}
