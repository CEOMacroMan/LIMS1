async function loadTable() {
  const debug = msg => {
    console.log(msg);
    const el = document.getElementById('debug');
    if (el) el.textContent += msg + '\n';
  };
  try {
    debug('Fetching TestData.xlsx...');
    const resp = await fetch('TestData.xlsx');
    debug(`Fetch status: ${resp.status}`);
    if (!resp.ok) throw new Error('unable to fetch TestData.xlsx');
    const buf = await resp.arrayBuffer();
    debug('Workbook loaded, parsing...');
    const wb = XLSX.read(buf, { type: 'array' });
    debug('Workbook sheets: ' + wb.SheetNames.join(', '));

    const name = wb.Workbook && wb.Workbook.Names
      ? wb.Workbook.Names.find(n => n.Name === 'INFOTable')
      : null;
    if (!name) throw new Error('table "INFOTable" not found');
    debug('Found table range: ' + name.Ref);

    const [sheetNameRaw, range] = name.Ref.split('!');
    const sheetName = sheetNameRaw.replace(/^'/, '').replace(/'$/, '');
    debug(`Using sheet "${sheetName}" range "${range}"`);
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`sheet "${sheetName}" not found`);
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range });
    debug('Rendering ' + rows.length + ' rows');
    render(rows);
  } catch (err) {
    debug('Error: ' + err.message);
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
