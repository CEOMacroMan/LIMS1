/* global ExcelJS */
export async function discoverTables(ab, sheetjsWb) {
  const ex = new ExcelJS.Workbook();
  await ex.xlsx.load(ab);
  const out = [];
  const sheets = [];
  ex.eachSheet(ws => {
    sheets.push({ sheet: ws.name, ref: ws.model && ws.model.ref });
    const list = [];
    if (typeof ws.eachTable === 'function') {
      ws.eachTable(t => list.push(t));
    } else if (ws.model && ws.model.tables) {
      Object.keys(ws.model.tables).forEach(n => {
        const t = ws.getTable ? ws.getTable(n) : null;
        if (t) list.push(t);
      });
    }
    list.forEach(tbl => {
      const name = tbl.name || tbl.displayName;
      const ref = tbl.tableRef || (tbl.model && tbl.model.tableRef) || (tbl.table && tbl.table.ref);
      if (name && ref) out.push({ type: 'table', sheet: ws.name, name, ref });
    });
  });
  const names = sheetjsWb && sheetjsWb.Workbook && sheetjsWb.Workbook.Names ? sheetjsWb.Workbook.Names : [];
  names.forEach(n => {
    if (!n.Ref) return;
    const parts = n.Ref.split('!');
    if (parts.length !== 2) return;
    const sheetName = parts[0].replace(/^'/, '').replace(/'$/, '');
    const ref = parts[1];
    out.push({ type: 'name', sheet: sheetName, name: n.Name, ref });
  });
  if (out.length === 0) {
    sheets.forEach(s => out.push({ type: 'sheet', sheet: s.sheet, name: s.sheet, ref: s.ref }));
  }
  return out;
}
