/* global XLSX */

function log(msg) {
  console.log(msg);
  const el = document.getElementById('debug');
  if (el) el.textContent += msg + '\n';
}

function clearLog() {
  const el = document.getElementById('debug');
  if (el) el.textContent = '';
}

function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = msg || '';
}

function renderTable(rows) {
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

function handleWorkbook(ab) {
  try {
    log('Parsing workbook...');
    const wb = XLSX.read(ab, { type: 'array' });

    const sheetNames = Object.keys(wb.Sheets || {});
    log('Detected sheets: ' + sheetNames.join(', '));

    const names = (wb.Workbook && Array.isArray(wb.Workbook.Names)) ? wb.Workbook.Names : [];
    if (names.length) {
      names.forEach(n => log(`Defined name: ${n.Name} -> ${n.Ref}`));
    } else {
      log('No defined names');
    }

    let tableSheetName;
    let tableRange;
    const infoName = names.find(n => n.Name === 'INFOTable');
    if (infoName) {
      log(`Found named range INFOTable: ${infoName.Ref}`);
      const idx = infoName.Ref.indexOf('!');
      if (idx !== -1) {
        const rawSheet = infoName.Ref.slice(0, idx);
        tableRange = infoName.Ref.slice(idx + 1);
        tableSheetName = rawSheet.replace(/^'/, '').replace(/'$/, '');
        log(`INFOTable sheet: ${tableSheetName}`);
        log(`INFOTable range: ${tableRange}`);
        const ws = wb.Sheets[tableSheetName];
        if (ws) {
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: tableRange });
          renderTable(rows);
          const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
          log(`Rendered ${rows.length} rows and ${colCount} columns`);
        } else {
          setStatus(`Sheet ${tableSheetName} not found`);
          log(`Sheet ${tableSheetName} not found`);
        }
      }
    } else {
      log("Named range 'INFOTable' not found");
      setStatus("Named range 'INFOTable' not found");
      document.getElementById('table').textContent = '';
    }

    const c1SheetName = tableSheetName || sheetNames[0];
    if (c1SheetName) {
      log(`Reading cell C1 from sheet ${c1SheetName}`);
      const ws = wb.Sheets[c1SheetName];
      if (ws) {
        const cell = ws['C1'];
        if (cell) {
          const val = cell.w || cell.v;
          document.getElementById('cellC1').textContent = 'C1: ' + val;
          log(`C1 value: ${val}`);
        } else {
          document.getElementById('cellC1').textContent = 'C1: (not found)';
          log('C1 not found');
        }
      } else {
        document.getElementById('cellC1').textContent = 'C1: (sheet not found)';
        log(`Sheet for C1 not found: ${c1SheetName}`);
      }
    } else {
      document.getElementById('cellC1').textContent = 'C1: (no sheets)';
      log('No sheets in workbook');
    }
  } catch (err) {
    log('Error parsing workbook: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error: ' + err.message);
  }
}

async function loadFromURL() {
  clearLog();
  setStatus('');
  document.getElementById('cellC1').textContent = '';
  document.getElementById('table').innerHTML = '';

  const url = document.getElementById('urlInput').value.trim();
  if (!url) {
    setStatus('Please enter a URL');
    log('No URL provided');
    return;
  }
  try {
    const resolved = new URL(url, window.location.href).href;
    log('Resolved URL: ' + resolved);
    const resp = await fetch(resolved);
    log(`Fetch status: ${resp.status} ${resp.statusText}`);
    log('content-type: ' + resp.headers.get('content-type'));
    log('content-length: ' + resp.headers.get('content-length'));
    if (!resp.ok) {
      const errMsg = `HTTP error ${resp.status} ${resp.statusText}`;
      log(errMsg);
      setStatus(errMsg);
      return;
    }
    const ab = await resp.arrayBuffer();
    log('ArrayBuffer length: ' + ab.byteLength);
    handleWorkbook(ab);
  } catch (err) {
    log('Error fetching URL: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error: ' + err.message);
  }
}

function loadFromFile() {
  clearLog();
  setStatus('');
  document.getElementById('cellC1').textContent = '';
  document.getElementById('table').innerHTML = '';

  const input = document.getElementById('fileInput');
  const file = input.files && input.files[0];
  if (!file) {
    setStatus('Please select a file');
    log('No file selected');
    return;
  }
  log(`Selected file: ${file.name} (${file.size} bytes)`);
  const reader = new FileReader();
  reader.onload = function (e) {
    const ab = e.target.result;
    log('File read: ' + ab.byteLength + ' bytes');
    handleWorkbook(ab);
  };
  reader.onerror = function (e) {
    const err = e.target.error;
    log('FileReader error: ' + (err && err.message));
    if (err && err.stack) log(err.stack);
    setStatus('Error reading file');
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById('loadUrlBtn').addEventListener('click', loadFromURL);
document.getElementById('loadFileBtn').addEventListener('click', loadFromFile);
