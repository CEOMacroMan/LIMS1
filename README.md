# INFOTable Viewer

Simple static page that reads the workbook `TestData.xls` and displays the table named `InfoTable` as an HTML table.

Serve the folder using any static file server:

```bash
npx serve .
```

Open the served URL, click **Choose File**, and select `TestData.xls` from the folder `C:\Hutchinson old`. Then click **Load** to render the table.

If loading fails, check the debug log under the table for step-by-step status messages showing each read and parse stage.
