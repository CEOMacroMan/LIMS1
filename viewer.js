/* global XLSX, ExcelJS, JSZip */


function log(msg) {
  console.log(msg);
  const el = document.getElementById('debug');
  if (el) el.textContent += msg + '\n';
}

const DEBUG_MODE = true;
function logKV(label, obj) {
  log(label + ': ' + (typeof obj === 'string' ? obj : JSON.stringify(obj)));
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
let originalFileAB = null; // original file ArrayBuffer for patching
const fsSupported = ('showOpenFilePicker' in window) && (location.protocol === 'https:' || location.hostname === 'localhost');

function updateSaveButton() {
  const btn = document.getElementById('saveBtn');
  if (btn) btn.disabled = !currentFileHandle;
  const fmtBtn = document.getElementById('saveFmtBtn');
  if (fmtBtn) fmtBtn.disabled = !currentFileHandle;
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

  const urlInput = document.getElementById('urlInput').value;
  const url = typeof urlInput === 'string' ? urlInput.trim() : '';
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
    originalFileAB = ab.slice(0);
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
    originalFileAB = ab.slice(0);
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
    originalFileAB = ab.slice(0);
    await handleWorkbook(ab);
    setStatus(`Loaded ${currentFileName} for editing`);
  } catch (err) {
    log('Open for edit error: ' + err.message);
    if (err.stack) log(err.stack);
    setStatus('Error opening file');
  }
  updateSaveButton();
}

function safeSample(v) {
  try {
    return typeof v === 'string' ? v.slice(0, 50) : JSON.stringify(v).slice(0, 120);
  } catch (e) {
    return String(v).slice(0, 120);
  }
}

function normalizeForWrite(raw) {
  if (raw == null || (typeof raw === 'string' && raw.trim() === '')) return { kind: 'blank' };
  if (typeof raw === 'number') {
    return isNaN(raw) ? { kind: 'string', v: String(raw) } : { kind: 'number', v: raw };
  }
  if (typeof raw === 'boolean') return { kind: 'boolean', v: raw };
  if (typeof raw === 'string') {
    const s = raw;
    const t = s.trim();
    if (t === '') return { kind: 'blank' };
    if (isFinite(Number(t))) return { kind: 'number', v: Number(t) };
    return { kind: 'string', v: s };
  }
  if (typeof raw === 'object') {
    if (raw && typeof raw.v !== 'undefined') return normalizeForWrite(raw.v);
    const prim = String(raw);
    const norm = normalizeForWrite(prim);
    if (norm.kind === 'string' && norm.v === prim) {
      log(`[normalizeForWrite] treating ${typeof raw} as string ${safeSample(prim)}`);
    }
    return norm;
  }
  return { kind: 'string', v: String(raw) };
}

function applyEdits() {
  if (!sheetjsWb || !currentSelection) return [];
  const ws = sheetjsWb.Sheets[currentSelection.sheet];
  if (!ws) return [];
  const errors = [];
  const patches = [];
  let processed = 0;
  for (let r = 0; r < editableData.length; ++r) {
    for (let c = 0; c < editableData[r].length; ++c) {
      const raw = editableData[r][c];
      const addr = XLSX.utils.encode_cell({ r: currentSelection.start.r + r, c: currentSelection.start.c + c });
      try {
        const norm = normalizeForWrite(raw);
        patches.push({ addr, norm });

        if (norm.kind === 'blank') {
          delete ws[addr];
        } else if (norm.kind === 'number') {
          ws[addr] = { t: 'n', v: norm.v };
        } else if (norm.kind === 'boolean') {
          ws[addr] = { t: 'b', v: norm.v };
        } else {
          ws[addr] = { t: 's', v: norm.v };
        }
      } catch (e) {
        errors.push({ r, c, addr, type: typeof raw, sample: safeSample(raw), message: e.message });
      }
      processed++;
      if (DEBUG_MODE && processed % 100 === 0) log(`[applyEdits] processed ${processed} cells`);
    }
  }
  if (errors.length) {
    log('[applyEdits] cell write errors: ' + JSON.stringify(errors.slice(0, 10)));
    const err = new Error('applyEdits failed for ' + errors.length + ' cells (see debug).');
    err.cells = errors;
    throw err;
  }
  return patches;
}

async function patchWorkbook(ab, patches, sheetName) {
  const zip = await JSZip.loadAsync(ab);
  const parser = new DOMParser();
  const serializer = new XMLSerializer();

  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbDoc = parser.parseFromString(wbXml, 'application/xml');
  const sheets = Array.from(wbDoc.getElementsByTagName('sheet'));
  let rId = null;
  sheets.forEach(s => {
    if (s.getAttribute('name') === sheetName) rId = s.getAttribute('r:id');
  });
  if (!rId) throw new Error('sheet mapping failed');

  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relsDoc = parser.parseFromString(relsXml, 'application/xml');
  const rels = Array.from(relsDoc.getElementsByTagName('Relationship'));
  let sheetPath = null;
  rels.forEach(r => {
    if (r.getAttribute('Id') === rId) sheetPath = 'xl/' + r.getAttribute('Target').replace(/^\//, '');
  });
  if (!sheetPath) throw new Error('worksheet path not found');

  const sheetXml = await zip.file(sheetPath).async('string');
  const sheetDoc = parser.parseFromString(sheetXml, 'application/xml');
  const sheetRoot = sheetDoc.documentElement;

  const sstPath = 'xl/sharedStrings.xml';
  let sstDoc;
  let shared = [];
  let sstCount = 0;
  if (zip.file(sstPath)) {
    const sstXml = await zip.file(sstPath).async('string');
    sstDoc = parser.parseFromString(sstXml, 'application/xml');
    const sis = Array.from(sstDoc.getElementsByTagName('si'));
    shared = sis.map(si => si.textContent);
    const root = sstDoc.documentElement;
    sstCount = Number(root.getAttribute('count')) || shared.length;
  } else {
    sstDoc = parser.parseFromString('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>', 'application/xml');
  }
  const sstRoot = sstDoc.documentElement;
  let sstUnique = shared.length;

  patches.forEach(p => {
    const cell = XLSX.utils.decode_cell(p.addr);
    const rowStr = String(cell.r + 1);
    const addr = p.addr;
    let rowNode = sheetDoc.querySelector(`row[r="${rowStr}"]`);
    if (!rowNode) {
      rowNode = sheetDoc.createElement('row');
      rowNode.setAttribute('r', rowStr);
      sheetRoot.appendChild(rowNode);
    }
    let cNode = rowNode.querySelector(`c[r="${addr}"]`);
    if (!cNode) {
      cNode = sheetDoc.createElement('c');
      cNode.setAttribute('r', addr);
      rowNode.appendChild(cNode);
    }
    const fNode = cNode.querySelector('f');
    if (fNode) cNode.removeChild(fNode);
    const vNode = cNode.querySelector('v');
    if (vNode) cNode.removeChild(vNode);
    cNode.removeAttribute('t');

    if (p.norm.kind === 'blank') {
      // nothing
    } else if (p.norm.kind === 'number') {
      cNode.setAttribute('t', 'n');
      const newV = sheetDoc.createElement('v');
      newV.textContent = String(p.norm.v);
      cNode.appendChild(newV);
    } else if (p.norm.kind === 'boolean') {
      cNode.setAttribute('t', 'b');
      const newV = sheetDoc.createElement('v');
      newV.textContent = p.norm.v ? '1' : '0';
      cNode.appendChild(newV);
    } else if (p.norm.kind === 'string') {
      cNode.setAttribute('t', 's');
      let idx = shared.indexOf(p.norm.v);
      if (idx === -1) {
        idx = shared.length;
        shared.push(p.norm.v);
        const si = sstDoc.createElement('si');
        const t = sstDoc.createElement('t');
        if (/^\s|\s$/.test(p.norm.v)) t.setAttribute('xml:space', 'preserve');
        t.textContent = p.norm.v;
        si.appendChild(t);
        sstRoot.appendChild(si);
        sstUnique++;
      }
      sstCount++;
      const newV = sheetDoc.createElement('v');
      newV.textContent = String(idx);
      cNode.appendChild(newV);
    }
  });
  sstRoot.setAttribute('count', String(sstCount));
  sstRoot.setAttribute('uniqueCount', String(sstUnique));
  zip.file(sheetPath, serializer.serializeToString(sheetDoc));
  zip.file(sstPath, serializer.serializeToString(sstDoc));
  return await zip.generateAsync({ type: 'arraybuffer' });

}

async function saveToOriginal() {
  if (!currentFileHandle) {
    setStatus('No editable file handle. Use "Open for edit (local)". Downloading copy...');
    log('Save: missing FileSystemFileHandle; using download copy');
    downloadCopy();
    return;
  }
  const selIdx = document.getElementById('tableList').selectedIndex;
  const info = tableEntries[selIdx] || {};
  logKV('[save] selection', {
    table: info.name || '',
    a1: info.ref || info.range || (currentSelection && currentSelection.range),
    sheet: currentSelection && currentSelection.sheet,
    start: currentSelection && currentSelection.start,
    end: currentSelection && currentSelection.end,
    rows: editableData.length,
    cols: Math.max(...editableData.map(r => r.length))
  });
  let step = 'start';

  try {
    step = 'getFile';
    await currentFileHandle.getFile();
    step = 'permission';
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
    step = 'applyEdits';
    const patches = applyEdits();
    step = 'build-binary';
    const bookType = currentBookType === 'xlsm' ? 'xlsm' : 'xlsx';
    const ab = XLSX.write(sheetjsWb, { type: 'array', bookType, bookVBA: currentBookType === 'xlsm' });
    step = 'write';

    const w = await currentFileHandle.createWritable();
    await w.write(ab);
    await w.close();
    originalFileAB = ab;
    log(`Saved via FS Access (${ab.byteLength} bytes)`);

    setStatus('File saved');
  } catch (err) {
    log('Save error: ' + err.message);
    logKV('[save] error', { action: 'save', step, message: err.message });

    logKV('[save] selection', {
      table: info.name || '',
      a1: info.ref || info.range || (currentSelection && currentSelection.range),
      sheet: currentSelection && currentSelection.sheet,
      start: currentSelection && currentSelection.start,
      end: currentSelection && currentSelection.end
    });
    if (err.cells) logKV('[save] cells', err.cells.slice(0, 10));
    if (err.stack) log(err.stack);
    setStatus('Error saving; using Download copy. Use "Open for edit (local)" to enable saving.');
    log('Save: falling back to download copy');
    downloadCopy();

  }
}

function downloadCopy() {
  if (!sheetjsWb) {
    setStatus('No workbook loaded');
    log('Download aborted: no workbook');
    return;
  }
  const selIdx = document.getElementById('tableList').selectedIndex;
  const info = tableEntries[selIdx] || {};
  logKV('[download] selection', {
    table: info.name || '',
    a1: info.ref || info.range || (currentSelection && currentSelection.range),
    sheet: currentSelection && currentSelection.sheet,
    start: currentSelection && currentSelection.start,
    end: currentSelection && currentSelection.end,
    rows: editableData.length,
    cols: Math.max(...editableData.map(r => r.length))
  });
  let step = 'start';
  try {
    step = 'applyEdits';
    const patches = applyEdits();
    step = 'build-binary';
    const bookType = currentBookType === 'xlsm' ? 'xlsm' : 'xlsx';
    const ab = XLSX.write(sheetjsWb, { type: 'array', bookType, bookVBA: currentBookType === 'xlsm' });
    originalFileAB = ab;
    step = 'download';

    const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, '') : 'workbook';
    const ext = currentBookType === 'xlsm' ? '.xlsm' : '.xlsx';
    const name = `${base}_edited${ext}`;
    XLSX.writeFile(sheetjsWb, name, { bookType, bookVBA: currentBookType === 'xlsm' });
    log(`Download copy initiated (${ab.byteLength} bytes)`);

    setStatus('Download started');
  } catch (err) {
    log('Download error: ' + err.message);
    logKV('[download] error', { action: 'download', step, message: err.message });

    logKV('[download] selection', {
      table: info.name || '',
      a1: info.ref || info.range || (currentSelection && currentSelection.range),
      sheet: currentSelection && currentSelection.sheet,
      start: currentSelection && currentSelection.start,
      end: currentSelection && currentSelection.end
    });
    if (err.cells) logKV('[download] cells', err.cells.slice(0, 10));
    if (err.stack) log(err.stack);
    setStatus('Error downloading file');
  }
}

async function saveToOriginalFmt() {
  if (!currentFileHandle) {
    setStatus('No editable file handle. Use "Open for edit (local)" first.');
    log('Save (preserve) aborted: no handle');
    return;
  }
  const selIdx = document.getElementById('tableList').selectedIndex;
  const info = tableEntries[selIdx] || {};
  logKV('[save-preserve] selection', {
    table: info.name || '',
    a1: info.ref || info.range || (currentSelection && currentSelection.range),
    sheet: currentSelection && currentSelection.sheet,
    start: currentSelection && currentSelection.start,
    end: currentSelection && currentSelection.end,
    rows: editableData.length,
    cols: Math.max(...editableData.map(r => r.length))
  });
  let step = 'start';
  try {
    step = 'getFile';
    const file = await currentFileHandle.getFile();
    const origAb = await file.arrayBuffer();
    step = 'permission';
    let perm = await currentFileHandle.queryPermission({ mode: 'readwrite' });
    log('Save(preserve): queryPermission -> ' + perm);
    if (perm !== 'granted') {
      perm = await currentFileHandle.requestPermission({ mode: 'readwrite' });
      log('Save(preserve): requestPermission -> ' + perm);
      if (perm !== 'granted') {
        setStatus('Write permission denied.');
        throw new Error('permission denied');
      }
    }
    step = 'applyEdits';
    const patches = applyEdits();
    step = 'build-binary';
    const patched = await patchWorkbook(origAb, patches, currentSelection.sheet);
    step = 'write';
    const w = await currentFileHandle.createWritable();
    await w.write(patched);
    await w.close();
    originalFileAB = patched;
    log(`Saved (preserve formatting) (${patched.byteLength} bytes)`);
    setStatus('File saved');
  } catch (err) {
    log('Save (preserve) error: ' + err.message);
    logKV('[save-preserve] error', { action: 'save-preserve', step, message: err.message });
    if (err.cells) logKV('[save-preserve] cells', err.cells.slice(0, 10));
    setStatus('Error saving file');
  }
}

async function downloadCopyFmt() {
  if (!originalFileAB) {
    setStatus('No workbook loaded');
    log('Download (preserve) aborted: no workbook');
    return;
  }
  const selIdx = document.getElementById('tableList').selectedIndex;
  const info = tableEntries[selIdx] || {};
  logKV('[download-preserve] selection', {
    table: info.name || '',
    a1: info.ref || info.range || (currentSelection && currentSelection.range),
    sheet: currentSelection && currentSelection.sheet,
    start: currentSelection && currentSelection.start,
    end: currentSelection && currentSelection.end,
    rows: editableData.length,
    cols: Math.max(...editableData.map(r => r.length))
  });
  let step = 'start';
  try {
    step = 'applyEdits';
    const patches = applyEdits();
    step = 'build-binary';
    const patched = await patchWorkbook(originalFileAB, patches, currentSelection.sheet);
    step = 'download';
    const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, '') : 'workbook';
    const ext = currentBookType === 'xlsm' ? '.xlsm' : '.xlsx';
    const name = `${base}_edited${ext}`;
    const blob = new Blob([patched], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    originalFileAB = patched;
    log(`Download (preserve) initiated (${patched.byteLength} bytes)`);
    setStatus('Download started');
  } catch (err) {
    log('Download (preserve) error: ' + err.message);
    logKV('[download-preserve] error', { action: 'download-preserve', step, message: err.message });
    if (err.cells) logKV('[download-preserve] cells', err.cells.slice(0, 10));
    log('Download (preserve) falling back to data-only copy');
    downloadCopy();
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
        currentSelection = { sheet: info.sheet, range, start: decoded.s, end: decoded.e };
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
      currentSelection = { sheet: info.sheet, range: info.ref, start: decoded.s, end: decoded.e };
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
      currentSelection = { sheet: info.sheet, range: rangeStr, start: range.s, end: range.e };
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
document.getElementById('saveFmtBtn').addEventListener('click', saveToOriginalFmt);

document.getElementById('downloadBtn').addEventListener('click', downloadCopy);
document.getElementById('downloadFmtBtn').addEventListener('click', downloadCopyFmt);

if (!fsSupported) {
  const msg = document.getElementById('fsMessage');
  if (msg) msg.textContent = 'FS Access API requires HTTPS or localhost. Use Download copy.';
  ['openFsBtn', 'saveBtn', 'saveFmtBtn'].forEach(id => {
    const b = document.getElementById(id);
    if (b) b.disabled = true;
  });
}

updateSaveButton();

/*
 * Usage:
 * Option A: open index.html and use "Load local file" (no CORS issues).
 * Option B: serve the Excel file over HTTP and use "Load via URL". Local
 *           paths like C:\... cannot be fetched by the browser; use
 *           http://localhost/... instead.
 */
