# INFOTable Viewer

Simple static page that reads the workbook `TestData.xlsx` and displays the table named `INFOTable` as an HTML table.

Place `TestData.xlsx` in the project root (for example, the OneDrive folder `C:\Users\gimenezherrerosergj\OneDrive - ApplusGlobal\CODEX\Applus Laboratories. IMA & Barcelona - PowerApps LIMS`) and serve the folder using any static file server:

```bash
npx serve .
```

Open the served URL, adjust the **Excel file path** field if necessary, and click **Load** to render the table.

If loading fails, check the debug log under the table for step-by-step status messages showing each fetch and parse stage.
