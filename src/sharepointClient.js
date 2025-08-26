// Thin client for SharePoint Excel access.
// Assumes authentication is already handled elsewhere.
// Uses Microsoft Graph table endpoints to fetch schema and sample rows.

export async function fetchTable(config) {
  const { sharepointSite, filePath, worksheet, table } = config;
  const base = `https://graph.microsoft.com/v1.0/sites/${sharepointSite}/drive/root:${filePath}:/workbook`;
  const ws = worksheet ? `/worksheets/${worksheet}` : '';
  const tablePath = `${base}${ws}/tables/${table}`;

  const [columnsResp, rowsResp] = await Promise.all([
    fetch(`${tablePath}/columns?$expand=dataValidation`),
    fetch(`${tablePath}/rows?$top=5`)
  ]);

  if (!columnsResp.ok) {
    throw new Error(`Spalten konnten nicht geladen werden (${columnsResp.status})`);
  }
  if (!rowsResp.ok) {
    throw new Error(`Zeilen konnten nicht geladen werden (${rowsResp.status})`);
  }

  const columnsJson = await columnsResp.json();
  const rowsJson = await rowsResp.json();

  const columns = columnsJson.value.map(col => ({
    name: col.name,
    address: col.address,
    validation: col.dataValidation,
  }));

  const rows = rowsJson.value.map(r => r.values[0]);

  return { columns, rows };
}
