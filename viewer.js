async function loadTable() {
  try {
    const resp = await fetch('TestData.xlsx');
    if (!resp.ok) throw new Error('unable to fetch TestData.xlsx');
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });

    const name = wb.Workbook && wb.Workbook.Names
      ? wb.Workbook.Names.find(n => n.Name === 'INFOTable')
      : null;
    if (!name) throw new Error('table "INFOTable" not found');

    const [sheetNameRaw, range] = name.Ref.split('!');
    const sheetName = sheetNameRaw.replace(/^'/, '').replace(/'$/, '');
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`sheet "${sheetName}" not found`);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range });
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
