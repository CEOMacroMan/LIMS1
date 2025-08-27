# INFOTable Viewer

Simple static page that reads the workbook `TestData.xlsm` and displays the table named `INFOTable` as an HTML table.

Serve the folder using any static file server:

```bash
npx serve .
```

Open the served URL and click **Load workbook**. The page fetches `TestData.xlsm` from the same directory, shows the value of cell `C1`, and then renders the `INFOTable` range.

If loading fails, check the debug log under the table for step-by-step status messages showing each read and parse stage.
