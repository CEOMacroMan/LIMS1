async function loadTable(file) {
  const debug = msg => {
    console.log(msg);
    const el = document.getElementById('debug');
    if (el) el.textContent += msg + '\n';
  };
  try {
    debug('Reading ' + file.name + '...');
    const buf = await file.arrayBuffer();
    debug('Workbook loaded, parsing...');
    const wb = XLSX.read(buf, { type: 'array' });
    debug('Workbook sheets: ' + wb.SheetNames.join(', '));

    const name = wb.Workbook?.Names
      ? wb.Workbook.Names.find(n => n.Name.toLowerCase() === 'infotable')
      : null;
    if (!name) throw new Error('table "InfoTable" not found');
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

document.getElementById('loadBtn').addEventListener('click', () => {
  const input = document.getElementById('fileInput');
  const file = input.files[0];
  document.getElementById('debug').textContent = '';
  if (!file) {
    const msg = 'No file selected';
    console.log(msg);
    document.getElementById('debug').textContent = msg;
    document.getElementById('table').textContent = msg;
    return;
  }
  document.getElementById('table').textContent = 'Loading...';
  loadTable(file);
});
