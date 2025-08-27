# INFOTable Viewer

Simple static page that reads the workbook `TestData.xlsx` and displays the table named `INFOTable` as an HTML table.

Place `TestData.xlsx` in the project root and serve the folder using any static file server:

```bash
npx serve .
```

Open the served URL and the table contents will be rendered in the browser.

If loading fails, check the debug log under the table for step-by-step status messages.
