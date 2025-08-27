# INFOTable Viewer
Simple static page that reads a workbook and displays the table named `INFOTable` as an HTML table.

Serve the folder using any static file server:

```bash
npx serve .
```

Open the served URL, enter the path to your workbook in the **Excel file path** field, and click **Load** to render the table. If loading fails, check the debug log under the table for step-by-step status messages showing each fetch and parse stage.
