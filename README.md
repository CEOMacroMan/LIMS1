# SharePoint Excel Form Demo

This project demonstrates loading an Excel table from SharePoint and generating a dynamic form based on the table schema. It is read-only and contains no data modification features.

## Quick start

```bash
# serve the static files
npx serve .
```

Open the served URL in a browser.

## Configuration

Edit `config.json` to point to your SharePoint file:

- `sharepointSite`: SharePoint site identifier.
- `filePath`: Path to the Excel file within the site.
- `worksheet` (optional): Worksheet name containing the table.
- `table`: Named table to read.
- `locale`: Locale string (e.g., `de-DE`).

Authentication is assumed to be handled externally.
