async function loadTable() {
  try {
    const resp = await fetch('INFOTable.xlsx');
    if (!resp.ok) throw new Error('unable to fetch INFOTable.xlsx');
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = 'INFOTable';
    let ws = wb.Sheets[sheetName];
    if (!ws) throw new Error('sheet "INFOTable" not found');
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    render(rows);
  } catch (err) {
    document.getElementById('table').textContent = 'Error: ' + err.message;
  }
}

function render(rows) {
  const container = document.getElementById('table');
  container.innerHTML = '';
  const table = document.createElement('table');
  rows.forEach(r => {
    const tr = document.createElement('tr');
    r.forEach(c => {
      const td = document.createElement('td');
      td.textContent = c !== undefined ? c : '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  container.appendChild(table);
}

loadTable();
