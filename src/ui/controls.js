import { clear } from '../utils/dom.js';

export function populateTableSelect(entries) {
  const sel = document.getElementById('tableList');
  clear(sel);
  entries.forEach((t, i) => {
    const opt = document.createElement('option');
    if (t.type === 'sheet') opt.textContent = `${t.sheet}: (preview first 50x50)`;
    else opt.textContent = `${t.sheet}: ${t.name} [${t.ref}]`;
    opt.value = i;
    sel.appendChild(opt);
  });
}

export function enableSave(enabled) {
  const b = document.getElementById('saveFmtBtn');
  if (b) b.disabled = !enabled;
}
