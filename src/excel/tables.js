import { logKV } from '../logger.js';

// Discover tables, named ranges, and sheets using ExcelJS
export async function discoverStructures(ab) {
  const entries = [];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ab);
  wb.eachSheet(ws => {
    const tables = ws.model && ws.model.tables ? ws.model.tables : [];
    tables.forEach(tbl => {
      entries.push({ type: 'table', sheet: ws.name, name: tbl.name, ref: tbl.tableRef });
    });
    entries.push({ type: 'sheet', sheet: ws.name });
  });
  const names = wb.definedNames && wb.definedNames.model && wb.definedNames.model.names ? wb.definedNames.model.names : [];
  names.forEach(n => {
    if (n.ranges && n.ranges.length) {
      const r = n.ranges[0];
      const sheet = wb.worksheets[r.sheetId] ? wb.worksheets[r.sheetId].name : r.sheetName || '';
      entries.push({ type: 'name', sheet, name: n.name, ref: r.range });
    }
  });
  logKV('[discover] entries', entries.length);
  return entries;
}
