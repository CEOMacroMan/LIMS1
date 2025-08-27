// Workbook expected alongside index.html, e.g., copied from
// C:\\Hutchinson old\\TestData.xlsm
const FILE_PATH = 'TestData.xlsm';

async function loadTable() {
  const debug = msg => {
    console.log(msg);
    const el = document.getElementById('debug');
    if (el) el.textContent += msg + '\n';
  };
  try {
    debug('Fetching ' + FILE_PATH + '...');
    const resp = await fetch(FILE_PATH);
    if (!resp.ok) throw new Error('unable to fetch ' + FILE_PATH);
    const buf = await resp.arrayBuffer();
    debug('Workbook loaded, parsing...');
    const wb = XLSX.read(buf, { type: 'array' });
    debug('Workbook sheets: ' + wb.SheetNames.join(', '));

    const name = wb.Workbook?.Names
      ? wb.Workbook.Names.find(n => n.Name.toLowerCase() === 'infotable')
      : null;
    if (!name) throw new Error('table "INFOTable" not found');
    debug('Found table range: ' + name.Ref);

    const [sheetNameRaw, range] = name.Ref.split('!');
    const sheetName = sheetNameRaw.replace(/^'/, '').replace(/'$/, '');
    debug(`Using sheet "${sheetName}" range "${range}"`);
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`sheet "${sheetName}" not found`);

    const c1 = ws['C1'] ? ws['C1'].v : '';
    debug('Cell C1 value: ' + c1);
    document.getElementById('cell').textContent = c1;

    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range });
    debug('Rendering ' + rows.length + ' rows');
    render(rows);
  } catch (err) {
    debug('Error: ' + err.message);
    document.getElementById('cell').textContent = 'Error';
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
  document.getElementById('debug').textContent = '';
  document.getElementById('cell').textContent = 'Loading...';
  document.getElementById('table').textContent = 'Loading...';
  loadTable();
});
