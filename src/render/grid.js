/* global XLSX */
import { state } from '../state.js';
import { clear } from '../utils/dom.js';
import { log } from '../logger.js';

let containerEl;

export function renderGrid(rows) {
  state.editableData = rows.map(r => r.slice());
  if (!containerEl) containerEl = document.getElementById('table');
  clear(containerEl);
  const table = document.createElement('table');
  const tbody = document.createElement('tbody');
  rows.forEach((row, rIdx) => {
    const tr = document.createElement('tr');
    for (let cIdx = 0; cIdx < row.length; ++cIdx) {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.className = 'cell';
      input.value = row[cIdx];
      input.dataset.r = rIdx;
      input.dataset.c = cIdx;
      td.appendChild(input);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  containerEl.appendChild(table);
}

function onInput(e) {
  const t = e.target;
  if (!t.classList.contains('cell')) return;
  const r = Number(t.dataset.r);
  const c = Number(t.dataset.c);
  state.editableData[r][c] = t.value;
  if (state.selection) {
    const addr = XLSX.utils.encode_cell({ r: state.selection.start.r + r, c: state.selection.start.c + c });
    log(`Edited ${addr} -> ${t.value}`);
  }
}

if (!containerEl) {
  containerEl = document.getElementById('table');
  containerEl.addEventListener('input', onInput);
}
