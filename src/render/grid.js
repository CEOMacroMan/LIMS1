import { state } from '../state.js';
import { clear, el } from '../utils/dom.js';

export function renderGrid(rows) {
  state.editableData = rows.map(r => r.map(c => (c == null ? '' : String(c))));
  const container = document.getElementById('table');
  clear(container);
  const table = el('table');
  state.editableData.forEach((row, r) => {
    const tr = el('tr');
    row.forEach((val, c) => {
      const td = el('td');
      const input = el('input', { className: 'cell', value: val });
      input.dataset.row = String(r);
      input.dataset.col = String(c);
      input.addEventListener('input', e => {
        const rr = Number(e.target.dataset.row);
        const cc = Number(e.target.dataset.col);
        state.editableData[rr][cc] = e.target.value;
      });
      td.appendChild(input);
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  container.appendChild(table);
}
