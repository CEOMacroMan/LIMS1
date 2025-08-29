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

function detectBookType(fname) {
  const lower = fname.toLowerCase();
  if (lower.endsWith('.xlsm')) return 'xlsm';
  if (lower.endsWith('.xlsb')) return 'xlsb';
  if (lower.endsWith('.xls')) return 'biff8';
  if (lower.endsWith('.ods')) return 'ods';
  return 'xlsx';
}

function mimeForBookType(bt) {
  switch (bt) {
    case 'xlsm':
      return 'application/vnd.ms-excel.sheet.macroEnabled.12';
    case 'xlsb':
      return 'application/vnd.ms-excel.sheet.binary.macroEnabled.12';
    case 'biff8':
      return 'application/vnd.ms-excel';
    case 'ods':
      return 'application/vnd.oasis.opendocument.spreadsheet';
    default:
      return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  }
}


function updateSaveButton() {
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
    currentBookType = detectBookType(currentFileName);

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
    currentBookType = detectBookType(file.name);

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

let currentFileLastModified = 0;

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
          'application/vnd.ms-excel': ['.xls', '.xlsb'],
          'application/vnd.oasis.opendocument.spreadsheet': ['.ods']
        }
      }]
    });
    currentFileHandle = handle;
    const file = await handle.getFile();
    currentFileName = file.name;
    currentBookType = detectBookType(file.name);

    currentFileLastModified = file.lastModified;
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
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  let step = 'load-zip';
  try {
    const zip = await JSZip.loadAsync(ab);
    step = 'workbook';
    const wbDoc = parser.parseFromString(await zip.file('xl/workbook.xml').async('string'), 'application/xml');
    const sheets = Array.from(wbDoc.getElementsByTagNameNS(NS, 'sheet'));
    let rId = null;
    sheets.forEach(s => {
      if (s.getAttribute('name') === sheetName) rId = s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    });
    if (!rId) {
      const err = new Error('sheet mapping failed');
      err.step = 'resolve-sheet';
      throw err;
    }
    step = 'rels';
    const relsDoc = parser.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
    const rels = Array.from(relsDoc.getElementsByTagNameNS(REL_NS, 'Relationship'));
    let sheetPath = null;
    rels.forEach(r => {
      if (r.getAttribute('Id') === rId) sheetPath = 'xl/' + r.getAttribute('Target').replace(/^\//, '');
    });
    if (!sheetPath) {
      const err = new Error('worksheet path not found');
      err.step = 'resolve-sheet';
      throw err;
    }
    step = 'sheet';
    const sheetDoc = parser.parseFromString(await zip.file(sheetPath).async('string'), 'application/xml');
    const sheetRoot = sheetDoc.documentElement;
    const sheetData = sheetRoot.getElementsByTagNameNS(NS, 'sheetData')[0];

    const sstPath = 'xl/sharedStrings.xml';
    const useSST = !!zip.file(sstPath);
    let sstDoc, sstRoot, shared = [], sstCount = 0, sstUnique = 0;
    if (useSST) {
      sstDoc = parser.parseFromString(await zip.file(sstPath).async('string'), 'application/xml');
      sstRoot = sstDoc.documentElement;
      const sis = Array.from(sstDoc.getElementsByTagNameNS(NS, 'si'));
      shared = sis.map(si => si.textContent);
      sstCount = Number(sstRoot.getAttribute('count')) || shared.length;
      sstUnique = Number(sstRoot.getAttribute('uniqueCount')) || shared.length;
    }

    const counts = { string: 0, number: 0, boolean: 0, blank: 0 };
    let sharedReused = 0, sharedAdded = 0;
    const xmlSpace = 'http://www.w3.org/XML/1998/namespace';

    patches.forEach(p => {
      const cell = XLSX.utils.decode_cell(p.addr);
      const rowStr = String(cell.r + 1);
      let rowNode = Array.from(sheetData.getElementsByTagNameNS(NS, 'row')).find(r => r.getAttribute('r') === rowStr);
      if (!rowNode) {
        rowNode = sheetDoc.createElementNS(NS, 'row');
        rowNode.setAttribute('r', rowStr);
        sheetData.appendChild(rowNode);
      }
      let cNode = Array.from(rowNode.getElementsByTagNameNS(NS, 'c')).find(c => c.getAttribute('r') === p.addr);
      if (!cNode) {
        cNode = sheetDoc.createElementNS(NS, 'c');
        cNode.setAttribute('r', p.addr);
        rowNode.appendChild(cNode);
      }
      const vNode = cNode.getElementsByTagNameNS(NS, 'v')[0];
      if (vNode) cNode.removeChild(vNode);
      const isNode = cNode.getElementsByTagNameNS(NS, 'is')[0];
      if (isNode) cNode.removeChild(isNode);
      const fNode = cNode.getElementsByTagNameNS(NS, 'f')[0];
      if (fNode) cNode.removeChild(fNode);
      cNode.removeAttribute('t');

      if (p.norm.kind === 'blank') {
        counts.blank++;
      } else if (p.norm.kind === 'number') {
        counts.number++;
        const newV = sheetDoc.createElementNS(NS, 'v');
        newV.textContent = String(p.norm.v);
        cNode.appendChild(newV);
      } else if (p.norm.kind === 'boolean') {
        counts.boolean++;
        cNode.setAttribute('t', 'b');
        const newV = sheetDoc.createElementNS(NS, 'v');
        newV.textContent = p.norm.v ? '1' : '0';
        cNode.appendChild(newV);
      } else if (p.norm.kind === 'string') {
        counts.string++;
        if (useSST) {
          cNode.setAttribute('t', 's');
          let idx = shared.indexOf(p.norm.v);
          if (idx === -1) {
            idx = shared.length;
            shared.push(p.norm.v);
            const si = sstDoc.createElementNS(NS, 'si');
            const t = sstDoc.createElementNS(NS, 't');
            if (/^\s|\s$/.test(p.norm.v)) t.setAttributeNS(xmlSpace, 'xml:space', 'preserve');
            t.textContent = p.norm.v;
            si.appendChild(t);
            sstRoot.appendChild(si);
            sstUnique++;
            sharedAdded++;
          } else {
            sharedReused++;
          }
          sstCount++;
          const newV = sheetDoc.createElementNS(NS, 'v');
          newV.textContent = String(idx);
          cNode.appendChild(newV);
        } else {
          cNode.setAttribute('t', 'inlineStr');
          const is = sheetDoc.createElementNS(NS, 'is');
          const t = sheetDoc.createElementNS(NS, 't');
          if (/^\s|\s$/.test(p.norm.v)) t.setAttributeNS(xmlSpace, 'xml:space', 'preserve');
          t.textContent = p.norm.v;
          is.appendChild(t);
          cNode.appendChild(is);
        }
      }
    });

    if (useSST) {
      sstRoot.setAttribute('count', String(sstCount));
      sstRoot.setAttribute('uniqueCount', String(sstUnique));
      zip.file(sstPath, serializer.serializeToString(sstDoc));
    }
    zip.file(sheetPath, serializer.serializeToString(sheetDoc));
    const abOut = await zip.generateAsync({ type: 'arraybuffer' });
    await JSZip.loadAsync(abOut); // validate zip
    return { ab: abOut, counts, shared: { reused: sharedReused, added: sharedAdded }, sheetPath, sstPath: useSST ? sstPath : null };
  } catch (err) {
    if (!err.step) err.step = step;
    throw err;
  }
}


async function verifyPatch(ab, patches, sheetPath, sstPath) {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const parser = new DOMParser();
  const zip = await JSZip.loadAsync(ab);
  const sheetDoc = parser.parseFromString(await zip.file(sheetPath).async('string'), 'application/xml');
  const sstDoc = sstPath && zip.file(sstPath) ? parser.parseFromString(await zip.file(sstPath).async('string'), 'application/xml') : null;
  const serializer = new XMLSerializer();
  const probes = [patches[0], patches[Math.floor(patches.length / 2)], patches[patches.length - 1]].filter(Boolean);
  for (const p of probes) {
    const cell = XLSX.utils.decode_cell(p.addr);
    const rowNode = Array.from(sheetDoc.getElementsByTagNameNS(NS, 'row')).find(r => r.getAttribute('r') === String(cell.r + 1));
    const cNode = rowNode && Array.from(rowNode.getElementsByTagNameNS(NS, 'c')).find(c => c.getAttribute('r') === p.addr);
    if (!cNode) {
      logKV('[verify-failed]', { addr: p.addr, reason: 'missing cell' });
      return false;
    }
    if (p.norm.kind === 'blank') {
      if (cNode.getElementsByTagNameNS(NS, 'v')[0] || cNode.getElementsByTagNameNS(NS, 'is')[0]) {
        logKV('[verify-failed]', { addr: p.addr, expected: 'blank', snippet: serializer.serializeToString(cNode) });
        return false;
      }
    } else if (p.norm.kind === 'number') {
      const vNode = cNode.getElementsByTagNameNS(NS, 'v')[0];
      if (!vNode || vNode.textContent !== String(p.norm.v)) {
        logKV('[verify-failed]', { addr: p.addr, expected: String(p.norm.v), snippet: serializer.serializeToString(cNode) });
        return false;
      }
    } else if (p.norm.kind === 'boolean') {
      const vNode = cNode.getElementsByTagNameNS(NS, 'v')[0];
      const expected = p.norm.v ? '1' : '0';
      if (!vNode || vNode.textContent !== expected) {
        logKV('[verify-failed]', { addr: p.addr, expected, snippet: serializer.serializeToString(cNode) });
        return false;
      }
    } else if (p.norm.kind === 'string') {
      if (sstDoc) {
        const vNode = cNode.getElementsByTagNameNS(NS, 'v')[0];
        if (!vNode) {
          logKV('[verify-failed]', { addr: p.addr, expected: p.norm.v });
          return false;
        }
        const idx = Number(vNode.textContent);
        const sis = sstDoc.getElementsByTagNameNS(NS, 'si');
        const text = sis[idx] && sis[idx].textContent;
        if (text !== p.norm.v) {
          logKV('[verify-failed]', { addr: p.addr, expected: p.norm.v, actual: text });
          return false;
        }
      } else {
        const isNode = cNode.getElementsByTagNameNS(NS, 'is')[0];
        const tNode = isNode && isNode.getElementsByTagNameNS(NS, 't')[0];
        const text = tNode && tNode.textContent;
        if (text !== p.norm.v) {
          logKV('[verify-failed]', { addr: p.addr, expected: p.norm.v, actual: text });
          return false;
        }

      }
      processed++;
      if (DEBUG_MODE && processed % 100 === 0) log(`[applyEdits] processed ${processed} cells`);
    }
  }
  return true;
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
    const bookType = currentBookType;
    const ab = XLSX.write(sheetjsWb, {
      type: 'array',
      bookType,
      bookVBA: ['xlsm', 'xlsb'].includes(bookType)
    });
    step = 'validate';
    await JSZip.loadAsync(ab);
    step = 'download';
    const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, '') : 'workbook';
    let ext = '.xlsx';
    if (bookType === 'xlsm') ext = '.xlsm';
    else if (bookType === 'xlsb') ext = '.xlsb';
    else if (bookType === 'biff8') ext = '.xls';
    else if (bookType === 'ods') ext = '.ods';
    const name = `${base}_edited${ext}`;
    XLSX.writeFile(sheetjsWb, name, { bookType, bookVBA: ['xlsm', 'xlsb'].includes(bookType) });
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
  if (!['xlsx', 'xlsm'].includes(currentBookType)) {
    setStatus('Format-preserving save only supports .xlsx/.xlsm');
    log('Save (preserve) aborted: unsupported book type ' + currentBookType);

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
    const freshFile = await currentFileHandle.getFile();
    log('Save(preserve): lastModified -> ' + freshFile.lastModified);
    if (currentFileLastModified && freshFile.lastModified !== currentFileLastModified) {
      if (!confirm('File changed on disk. Overwrite?')) {
        const err = new Error('external modification');
        err.step = 'external-change';
        throw err;
      }
    }
    const origAb = await freshFile.arrayBuffer();

    step = 'permission';
    let perm = await currentFileHandle.queryPermission({ mode: 'readwrite' });
    log('Save(preserve): queryPermission -> ' + perm);
    if (perm !== 'granted') {
      perm = await currentFileHandle.requestPermission({ mode: 'readwrite' });
      log('Save(preserve): requestPermission -> ' + perm);
      if (perm !== 'granted') {
        const err = new Error('permission denied');
        err.step = 'permission';
        throw err;
      }
    }
    step = 'applyEdits';
    const patches = applyEdits();
    step = 'build';
    let patchedInfo;
    try {
      patchedInfo = await patchWorkbook(origAb, patches, currentSelection.sheet);
    } catch (e) {
      e.step = e.step || 'build';
      throw e;
    }
    step = 'verify';
    const ok = await verifyPatch(patchedInfo.ab, patches, patchedInfo.sheetPath, patchedInfo.sstPath);
    if (!ok) {
      const err = new Error('verification failed');
      err.step = 'verify';
      throw err;
    }
    step = 'write';
    const w = await currentFileHandle.createWritable();
    await w.truncate(0);
    await w.write(new Uint8Array(patchedInfo.ab));

    await w.close();
    const after = await currentFileHandle.getFile();
    currentFileLastModified = after.lastModified;
    originalFileAB = patchedInfo.ab;
    logKV('[save-preserve] counts', {
      strings: patchedInfo.counts.string,
      numbers: patchedInfo.counts.number,
      booleans: patchedInfo.counts.boolean,
      blanks: patchedInfo.counts.blank,
      sharedReused: patchedInfo.shared.reused,
      sharedAdded: patchedInfo.shared.added
    });
    log(`Saved (preserve formatting) (${patchedInfo.ab.byteLength} bytes)`);
    setStatus('File saved');
  } catch (err) {
    log('Save (preserve) error: ' + err.message);
    logKV('[save-preserve] error', { action: 'save-preserve', step: err.step || step, message: err.message, name: err.name, stack: err.stack });
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
  if (!['xlsx', 'xlsm'].includes(currentBookType)) {
    setStatus('Format-preserving download only supports .xlsx/.xlsm');
    log('Download (preserve) aborted: unsupported book type ' + currentBookType);
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
    step = 'build';
    let patchedInfo;
    try {
      patchedInfo = await patchWorkbook(originalFileAB, patches, currentSelection.sheet);
    } catch (e) {
      e.step = e.step || 'build';
      throw e;
    }
    step = 'verify';
    const ok = await verifyPatch(patchedInfo.ab, patches, patchedInfo.sheetPath, patchedInfo.sstPath);
    if (!ok) {
      const err = new Error('verification failed');
      err.step = 'verify';
      throw err;
    }
    step = 'download';
    const base = currentFileName ? currentFileName.replace(/\.[^.]+$/, '') : 'workbook';
    let ext = '.xlsx';
    if (currentBookType === 'xlsm') ext = '.xlsm';
    else if (currentBookType === 'xlsb') ext = '.xlsb';
    else if (currentBookType === 'biff8') ext = '.xls';
    else if (currentBookType === 'ods') ext = '.ods';
    const name = `${base}_edited${ext}`;
    const blob = new Blob([patchedInfo.ab], { type: mimeForBookType(currentBookType) });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    originalFileAB = patchedInfo.ab;
    logKV('[download-preserve] counts', {
      strings: patchedInfo.counts.string,
      numbers: patchedInfo.counts.number,
      booleans: patchedInfo.counts.boolean,
      blanks: patchedInfo.counts.blank,
      sharedReused: patchedInfo.shared.reused,
      sharedAdded: patchedInfo.shared.added
    });
    log(`Download (preserve) initiated (${patchedInfo.ab.byteLength} bytes)`);
    setStatus('Download started');
  } catch (err) {
    log('Download (preserve) error: ' + err.message);
    logKV('[download-preserve] error', { action: 'download-preserve', step: err.step || step, message: err.message, name: err.name, stack: err.stack });
    if (err.cells) logKV('[download-preserve] cells', err.cells.slice(0, 10));
    setStatus('Error building download');

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
document.getElementById('saveFmtBtn').addEventListener('click', saveToOriginalFmt);

document.getElementById('downloadBtn').addEventListener('click', downloadCopy);
document.getElementById('downloadFmtBtn').addEventListener('click', downloadCopyFmt);

if (!fsSupported) {
  const msg = document.getElementById('fsMessage');
  if (msg) msg.textContent = 'FS Access API requires HTTPS or localhost. Use Download copy.';
  ['openFsBtn', 'saveFmtBtn'].forEach(id => {

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
