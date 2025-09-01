/* global JSZip, DOMParser, XMLSerializer, XLSX */
import { state } from '../state.js';
import { log, logKV } from '../logger.js';

export async function buildPreserveBinary() {
  const { originalAb, selection, workbook, originalExt } = state;
  let zip;
  try {
    zip = await JSZip.loadAsync(originalAb);
  } catch (err) {
    err.step = 'zip-validate';
    throw err;
  }
  let sheetPath;
  const parser = new DOMParser();
  try {
    const wbXml = await zip.file('xl/workbook.xml').async('text');
    const wbDoc = parser.parseFromString(wbXml, 'application/xml');
    const relXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');
    const relDoc = parser.parseFromString(relXml, 'application/xml');
    const sheets = wbDoc.getElementsByTagName('sheet');
    let rId;
    for (let i = 0; i < sheets.length; ++i) {
      if (sheets[i].getAttribute('name') === selection.sheet) {
        rId = sheets[i].getAttribute('r:id');
        break;
      }
    }
    if (!rId) throw new Error('sheet not found');
    const rels = relDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < rels.length; ++i) {
      if (rels[i].getAttribute('Id') === rId) {
        const tgt = rels[i].getAttribute('Target');
        sheetPath = 'xl/' + tgt.replace(/^\.\//, '');
        break;
      }
    }
    if (!sheetPath) throw new Error('sheet path not found');
  } catch (err) {
    err.step = 'resolve-sheet';
    throw err;
  }
  let sheetDoc;
  try {
    const sheetXml = await zip.file(sheetPath).async('text');
    sheetDoc = parser.parseFromString(sheetXml, 'application/xml');
  } catch (err) {
    err.step = 'parse-xml';
    throw err;
  }
  let sstDoc = null;
  let sstPath = 'xl/sharedStrings.xml';
  let sstList = [];
  let sstChanged = false;
  if (zip.file(sstPath)) {
    try {
      const sstXml = await zip.file(sstPath).async('text');
      sstDoc = parser.parseFromString(sstXml, 'application/xml');
      const sis = sstDoc.getElementsByTagName('si');
      for (let i = 0; i < sis.length; ++i) {
        sstList.push(sis[i].textContent);
      }
    } catch (err) {
      err.step = 'sharedStrings';
      throw err;
    }
  }
  function getSstIndex(str) {
    const idx = sstList.indexOf(str);
    if (idx !== -1) return idx;
    sstList.push(str);
    const si = sstDoc.createElement('si');
    const t = sstDoc.createElement('t');
    t.textContent = str;
    si.appendChild(t);
    sstDoc.documentElement.appendChild(si);
    const count = Number(sstDoc.documentElement.getAttribute('count') || '0') + 1;
    sstDoc.documentElement.setAttribute('count', String(count));
    const unique = Number(sstDoc.documentElement.getAttribute('uniqueCount') || '0') + 1;
    sstDoc.documentElement.setAttribute('uniqueCount', String(unique));
    sstChanged = true;
    return sstList.length - 1;
  }
  try {
    const sheetData = sheetDoc.getElementsByTagName('sheetData')[0];
    const start = selection.start;
    let processed = 0;
    for (let r = 0; r < state.editableData.length; ++r) {
      const rowData = state.editableData[r];
      const rnum = start.r + r + 1;
      let rowEl = sheetData.querySelector(`row[r="${rnum}"]`);
      if (!rowEl) {
        rowEl = sheetDoc.createElement('row');
        rowEl.setAttribute('r', String(rnum));
        sheetData.appendChild(rowEl);
      }
      for (let c = 0; c < rowData.length; ++c) {
        const addr = XLSX.utils.encode_cell({ r: start.r + r, c: start.c + c });
        const cellVal = workbook.Sheets[selection.sheet][addr];
        let cEl = rowEl.querySelector(`c[r="${addr}"]`);
        if (!cEl) {
          cEl = sheetDoc.createElement('c');
          cEl.setAttribute('r', addr);
          rowEl.appendChild(cEl);
        }
        Array.from(cEl.querySelectorAll('v,is')).forEach(n => cEl.removeChild(n));
        cEl.removeAttribute('t');
        if (!cellVal || cellVal.v === undefined || cellVal.v === null) {
          // blank
        } else if (cellVal.t === 'n') {
          const v = sheetDoc.createElement('v');
          v.textContent = String(cellVal.v);
          cEl.appendChild(v);
        } else if (cellVal.t === 'b') {
          cEl.setAttribute('t', 'b');
          const v = sheetDoc.createElement('v');
          v.textContent = cellVal.v ? '1' : '0';
          cEl.appendChild(v);
        } else if (cellVal.t === 's') {
          if (sstDoc) {
            const idx = getSstIndex(String(cellVal.v));
            cEl.setAttribute('t', 's');
            const v = sheetDoc.createElement('v');
            v.textContent = String(idx);
            cEl.appendChild(v);
          } else {
            cEl.setAttribute('t', 'inlineStr');
            const is = sheetDoc.createElement('is');
            const t = sheetDoc.createElement('t');
            t.textContent = String(cellVal.v);
            is.appendChild(t);
            cEl.appendChild(is);
          }
        }
        processed++;
        if (processed % 100 === 0) log(`[zip] processed ${processed} cells`);
      }
    }
    const serializer = new XMLSerializer();
    zip.file(sheetPath, serializer.serializeToString(sheetDoc));
    if (sstDoc && sstChanged) {
      zip.file(sstPath, serializer.serializeToString(sstDoc));
    }
  } catch (err) {
    err.step = err.step || 'write-cell';
    throw err;
  }
  let outAb;
  try {
    outAb = await zip.generateAsync({ type: 'arraybuffer' });
    logKV('[out-binary]', { ext: originalExt, byteLength: outAb.byteLength });
  } catch (err) {
    err.step = 'repack';
    throw err;
  }
  try {
    await JSZip.loadAsync(outAb);
    log('[zip-validate] ok');
  } catch (err) {
    err.step = 'zip-validate';
    throw err;
  }
  return outAb;
}
