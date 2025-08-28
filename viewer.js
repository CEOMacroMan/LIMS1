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
let editableData = []; // 2D array holding grid values
let currentSelection = null; // metadata for selected range
let currentFileHandle = null; // FileSystemFileHandle when opened via FS access
let currentFileName = '';
let currentBookType = 'xlsx';
const fsSupported = ('showOpenFilePicker' in window);

function updateSaveButton() {
  const btn = document.getElementById('saveBtn');
  if (btn) btn.disabled = !currentFileHandle;
}

function renderTable(rows) {
  editableData = rows.map(r => r.slice());
  const container = document.getElementById('table');
  container.innerHTML = '';
  const table = document.createElement('table');
  editableData.forEach((r, rIdx) => {
    const tr = document.createElement('tr');
    r.forEach((c, cIdx) => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.className = 'cell';
      input.value = c;
      input.dataset.row = String(rIdx);
      input.dataset.col = String(cIdx);
      input.addEventListener('input', e => {
        const row = Number(e.target.dataset.row);
        const col = Number(e.target.dataset.col);
        editableData[row][col] = e.target.value;
        if (currentSelection) {
          const addr = XLSX.utils.encode_cell({ r: currentSelection.start.r + row, c: currentSelection.start.c + col });
          log(`Edited ${addr} -> ${e.target.value}`);
        }
      });
      td.appendChild(input);
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
    const wb = XLSX.read(ab, { type: 'array', bookVBA: true });
    sheetjsWb = wb;

    const sheetNames = Object.keys(wb.Sheets || {});
    log('Detected sheets: ' + sheetNames.join(', '));

    const names = (wb.Workbook && Array.isArray(wb.Workbook.Names)) ? wb.Workbook.Names : [];
    if (names.length) {
      names.forEach(n => log(`Defined name: ${n.Name} -> ${n.Ref}`));
    } else {
      log('No defined names');
    }
    const infoName = names.find(n => n.Name === 'INFOTable');
    if (infoName) {
      const idx = infoName.Ref.indexOf('!');
      const sheet = idx !== -1 ? infoName.Ref.slice(0, idx).replace(/^'/, '').replace(/'$/, '') : '(unknown)';
      const ref = idx !== -1 ? infoName.Ref.slice(idx + 1) : infoName.Ref;
      log(`Found named range 'INFOTable' on sheet ${sheet} range ${ref}`);
    } else {

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
          const raw = t && (t.tableRef || (t.table && t.table.tableRef) || t.ref || '');
          const addr = raw.includes('!') ? raw.split('!')[1] : raw;

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
    currentFileHandle = null;
    currentFileName = resolved.split('/').pop().split('?')[0] || 'workbook.xlsx';
    currentBookType = currentFileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx';
    await handleWorkbook(ab);
    setStatus(`Loaded ${currentFileName}`);
    updateSaveButton();
  } catch (err) {
    log('Error fetching URL: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error: ' + err.message);
  }
}

async function loadFromFile() {

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
  reader.onload = async function (e) {
    const ab = e.target.result;
    log('File read: ' + ab.byteLength + ' bytes');
    currentFileHandle = null;
    currentFileName = file.name;
    currentBookType = file.name.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx';
    await handleWorkbook(ab);
    setStatus(`Loaded ${file.name} (use "Open for edit" to enable saving)`);
    updateSaveButton();
  };
  reader.onerror = function (e) {
    const err = e.target.error;
    log('FileReader error: ' + (err && err.message));
    if (err && err.stack) log(err.stack);
    setStatus('Error reading file');
  };
  reader.readAsArrayBuffer(file);
}

async function openForEdit() {
  if (!fsSupported) {
    setStatus('File System Access API not supported');
    log('FS access not supported');
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{
        description: 'Excel Files',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'],
          'application/vnd.ms-excel': ['.xls', '.xlsb']
        }
      }]
    });
    currentFileHandle = handle;
    const file = await handle.getFile();
    currentFileName = file.name;
    currentBookType = currentFileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx';
    const ab = await file.arrayBuffer();
    log(`Opened ${currentFileName} via FS Access (${ab.byteLength} bytes)`);
    await handleWorkbook(ab);
    setStatus(`Loaded ${currentFileName} for editing`);
  } catch (err) {
    log('Open for edit error: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error opening file');
  }
  updateSaveButton();
}

function applyEdits() {
  if (!sheetjsWb || !currentSelection) return;
  const ws = sheetjsWb.Sheets[currentSelection.sheet];
  if (!ws) return;
  for (let r = 0; r < editableData.length; ++r) {
    for (let c = 0; c < editableData[r].length; ++c) {
      const val = editableData[r][c];
      const addr = XLSX.utils.encode_cell({ r: currentSelection.start.r + r, c: currentSelection.start.c + c });
      if (val === '') {
        delete ws[addr];
      } else if (val.trim() !== '' && isFinite(Number(val))) {
        ws[addr] = { t: 'n', v: Number(val) };
      } else {
        ws[addr] = { t: 's', v: val };
      }
    }
  }
}

async function saveToOriginal() {
  if (!currentFileHandle) {
    setStatus('No editable file handle. Use "Open for edit (local)". Downloading copy...');
    log('Save: missing FileSystemFileHandle; using download copy');
    downloadCopy();
    return;
  }
  try {
    await currentFileHandle.getFile();
    let perm = await currentFileHandle.queryPermission({ mode: 'readwrite' });
    log('Save: queryPermission -> ' + perm);
    if (perm !== 'granted') {
      perm = await currentFileHandle.requestPermission({ mode: 'readwrite' });
      log('Save: requestPermission -> ' + perm);
      if (perm !== 'granted') {
        setStatus('Write permission denied. Downloading copy.');
        log('Save: permission denied; using download copy');
        downloadCopy();
        return;
      }
    }
    applyEdits();
    const ab = XLSX.write(sheetjsWb, { bookType: currentBookType, type: 'array', bookVBA: true });
    const w = await currentFileHandle.createWritable();
    await w.write(ab);
    await w.close();
    log(`Saved via FS Access (${ab.byteLength} bytes)`);
    setStatus('File saved');
  } catch (err) {
    log('Save error: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error saving; using Download copy. Use "Open for edit (local)" to enable saving.');
    downloadCopy();
  }
}

function downloadCopy() {
  if (!sheetjsWb) {
    setStatus('No workbook loaded');
    log('Download aborted: no workbook');
    return;
  }
  try {
    applyEdits();
    const ab = XLSX.write(sheetjsWb, { bookType: currentBookType, type: 'array', bookVBA: true });
    const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, '') : 'workbook';
    const ext = currentBookType === 'xlsm' ? '.xlsm' : '.xlsx';
    const name = `${base}_edited${ext}`;
    XLSX.writeFile(sheetjsWb, name, { bookType: currentBookType, bookVBA: true });
    log(`Download copy initiated (${ab.byteLength} bytes)`);
    setStatus('Download started');
  } catch (err) {
    log('Download error: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error downloading file');
  }
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
  if (info.type === 'table') {
    log(`Rendering table '${info.name}' range ${info.ref}`);
    const ws = sheetjsWb.Sheets[info.sheet];
    if (ws) {
      const range = info.ref;
      if (range) {
        const decoded = XLSX.utils.decode_range(range);
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range, defval: '' });
        currentSelection = { sheet: info.sheet, range, start: decoded.s };
        renderTable(rows);
        const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
        log(`Rendered ${rows.length} rows and ${colCount} columns`);
      } else {
        log(`Invalid table range for ${info.name}: ${info.ref}`);
        setStatus(`Invalid range ${info.ref}`);
      }
    } else {
      log(`Sheet ${info.sheet} not found`);
      setStatus(`Sheet ${info.sheet} not found`);
    }
  } else if (info.type === 'name') {
    log(`Rendering named range ${info.name} on sheet ${info.sheet} range ${info.ref}`);
    const ws = sheetjsWb.Sheets[info.sheet];
    if (ws) {
      const decoded = XLSX.utils.decode_range(info.ref);
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: info.ref, defval: '' });
      currentSelection = { sheet: info.sheet, range: info.ref, start: decoded.s };
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
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, range: rangeStr, defval: '' });
      currentSelection = { sheet: info.sheet, range: rangeStr, start: range.s };
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
document.getElementById('openFsBtn').addEventListener('click', openForEdit);
document.getElementById('saveBtn').addEventListener('click', saveToOriginal);
document.getElementById('downloadBtn').addEventListener('click', downloadCopy);

if (!fsSupported) {
  const msg = document.getElementById('fsMessage');
  if (msg) msg.textContent = 'FS Access API not supported; use Download copy.';
  const openBtn = document.getElementById('openFsBtn');
  if (openBtn) openBtn.disabled = true;
}

updateSaveButton();

/*
 * Usage:
 * Option A: open index.html and use "Load local file" (no CORS issues).
 * Option B: serve the Excel file over HTTP and use "Load via URL". Local
 *           paths like C:\... cannot be fetched by the browser; use
 *           http://localhost/... instead.
 */
