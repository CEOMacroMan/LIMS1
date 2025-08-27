/* global XLSX, ExcelJS */

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

let sheetjsWb; // SheetJS workbook
let tableEntries = []; // options for dropdown

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

function populateDropdown() {
  const sel = document.getElementById('tableList');
  sel.innerHTML = '';
  tableEntries.forEach((t, i) => {
    const opt = document.createElement('option');
    if (t.type === 'table') {
      opt.textContent = `${t.sheet}: ${t.name} [${t.ref}]`;
    } else if (t.type === 'name') {
      opt.textContent = `${t.sheet}: ${t.name} [${t.ref}]`;
    } else {
      opt.textContent = `${t.sheet}: (preview first 50x50)`;
    }
    opt.value = String(i);
    sel.appendChild(opt);
  });
}

function showC1(sheetName) {
  const cellDiv = document.getElementById('cellC1');
  if (!sheetjsWb) {
    cellDiv.textContent = 'C1: (no workbook)';
    return;
  }
  const sheetNames = Object.keys(sheetjsWb.Sheets || {});
  let target = sheetName;
  if (!target) {
    target = sheetNames.includes('INFO') ? 'INFO' : sheetNames[0];
  }
  if (!target) {
    cellDiv.textContent = 'C1: (no sheets)';
    log('No sheets in workbook');
    return;
  }
  log(`Reading cell C1 from sheet ${target}`);
  const ws = sheetjsWb.Sheets[target];
  if (!ws) {
    cellDiv.textContent = 'C1: (sheet not found)';
    log(`Sheet for C1 not found: ${target}`);
    return;
  }
  const cell = ws['C1'];
  if (cell) {
    const val = cell.w || cell.v;
    cellDiv.textContent = 'C1: ' + val;
    log(`C1 value: ${val}`);
  } else {
    cellDiv.textContent = 'C1: (not found)';
    log('C1 not found');
  }
}

async function handleWorkbook(ab) {
  try {
    log('Parsing workbook with SheetJS...');
    const wb = XLSX.read(ab, { type: 'array' });
    sheetjsWb = wb;

    const sheetNames = Object.keys(wb.Sheets || {});
    log('Detected sheets: ' + sheetNames.join(', '));

    const names = (wb.Workbook && Array.isArray(wb.Workbook.Names)) ? wb.Workbook.Names : [];
    if (names.length) {
      names.forEach(n => log(`Defined name: ${n.Name} -> ${n.Ref}`));
    } else {
      log('No defined names');
    }
    if (!names.find(n => n.Name === 'INFOTable')) {
      log("Named range 'INFOTable' not found");
    }

    tableEntries = [];

    if (typeof ExcelJS !== 'undefined') {
      log('Loading workbook with ExcelJS for table discovery...');
      const ex = new ExcelJS.Workbook();
      await ex.xlsx.load(ab);
      ex.eachSheet(ws => {
        const wsTables = ws.tables || {};
        Object.keys(wsTables).forEach(tname => {
          const t = ws.getTable ? ws.getTable(tname) : wsTables[tname];
          const addr = t && t.table ? t.table.ref : (t && t.ref ? t.ref : '');
          tableEntries.push({ type: 'table', sheet: ws.name, name: tname, ref: addr });
          log(`Table: ${tname} on sheet ${ws.name} range ${addr}`);
        });
      });
    } else {
      log('ExcelJS not available');
    }

    if (!tableEntries.length) {
      log('No Excel Tables found');
      if (names.length) {
        names.forEach(n => {
          const idx = n.Ref.indexOf('!');
          let sheet = sheetNames[0];
          let ref = n.Ref;
          if (idx !== -1) {
            sheet = n.Ref.slice(0, idx).replace(/^'/, '').replace(/'$/, '');
            ref = n.Ref.slice(idx + 1);
          }
          tableEntries.push({ type: 'name', sheet, name: n.Name, ref });
        });
      } else {
        sheetNames.forEach(sn => {
          tableEntries.push({ type: 'sheet', sheet: sn, name: sn });
        });
      }
    }

    populateDropdown();
    const firstSheet = tableEntries[0] ? tableEntries[0].sheet : (sheetNames.includes('INFO') ? 'INFO' : sheetNames[0]);
    showC1(firstSheet);
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
    await handleWorkbook(ab);
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

  function renderSelected() {
    setStatus('');
    const sel = document.getElementById('tableList');
    const idx = sel.selectedIndex;
    const info = tableEntries[idx];
    if (!info) {
      log('No selection to render');
      setStatus('Please select a table or range');
      return;
    }
    if (!sheetjsWb) {
      log('No workbook loaded');
      setStatus('No workbook loaded');
      return;
    }
    if (info.type === 'table' || info.type === 'name') {
      log(`Rendering ${info.type === 'table' ? 'table' : 'named range'} ${info.name} on sheet ${info.sheet} range ${info.ref}`);
      const ws = sheetjsWb.Sheets[info.sheet];
      if (ws) {
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: info.ref });
        renderTable(rows);
        const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
        log(`Rendered ${rows.length} rows and ${colCount} columns`);
      } else {
        log(`Sheet ${info.sheet} not found`);
        setStatus(`Sheet ${info.sheet} not found`);
      }
    } else if (info.type === 'sheet') {
      log(`Rendering sheet preview for ${info.sheet}`);
      const ws = sheetjsWb.Sheets[info.sheet];
      if (ws) {
        const used = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        const range = {
          s: { r: used.s.r, c: used.s.c },
          e: { r: Math.min(used.e.r, used.s.r + 49), c: Math.min(used.e.c, used.s.c + 49) }
        };
        const rangeStr = XLSX.utils.encode_range(range);
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: rangeStr });
        renderTable(rows);
        const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
        log(`Rendered preview ${rows.length} rows and ${colCount} columns from ${info.sheet} (${rangeStr})`);
      } else {
        log(`Sheet ${info.sheet} not found`);
        setStatus(`Sheet ${info.sheet} not found`);
      }
    }
    showC1(info.sheet);
  }

  document.getElementById('loadUrlBtn').addEventListener('click', loadFromURL);
  document.getElementById('loadFileBtn').addEventListener('click', loadFromFile);
  document.getElementById('renderBtn').addEventListener('click', renderSelected);

  /*
   * Usage:
   * Option A: open index.html and use "Load local file" (no CORS issues).
   * Option B: serve the Excel file over HTTP and use "Load via URL". Local
   *           paths like C:\... cannot be fetched by the browser; use
   *           http://localhost/... instead.
   */

