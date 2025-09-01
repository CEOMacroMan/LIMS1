/* global ExcelJS */
export async function discoverTables(ab) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ab);
  const out = [];
  wb.eachSheet(ws => {
    const tbls = ws.model && ws.model.tables ? ws.model.tables : {};
    Object.values(tbls).forEach(tbl => {
      const name = tbl.name || tbl.displayName || (tbl.table && tbl.table.name);
      const ref = tbl.ref || (tbl.table && tbl.table.ref);
      out.push({ type: 'table', sheet: ws.name, name, ref });
    });
  });
  const names = wb.definedNames && wb.definedNames.model && wb.definedNames.model.names ? wb.definedNames.model.names : [];
  names.forEach(n => {
    const first = n.ranges && n.ranges[0];
    if (first) {
      const [sheet, ref] = first.split('!');
      const sheetName = sheet.replace(/^'/, '').replace(/'$/, '');
      out.push({ type: 'name', sheet: sheetName, name: n.name, ref });
    }
  });
  return out;
}
