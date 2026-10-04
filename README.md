# Clear Sheet

A read-only Windows 10/11 x64 spreadsheet viewer built with Electron, React, TypeScript, Vite, MUI, SheetJS, AG Grid Community, and Zustand.

## Run and package

Use Node.js 22 or newer on Windows:

```sh
npm install
npm run dev
npm test
npm run dist
```

`npm start` runs the production build. The NSIS installer is generated in `release/`. The installer is unsigned; a signing certificate is needed for trusted publisher distribution.

## Features

- Native Open dialog, Ctrl+O, drag/drop, startup file arguments and Windows file associations.
- XLSX, XLS, XLSM, XLSB and CSV; multiple sheet tabs, loaded on demand.
- Virtualized rows and columns, resizable columns, original spreadsheet row/column addresses.
- Column sorting and text filters; debounced search filters matching rows in the active sheet.
- Zoom, formula bar, focused-cell copy with Ctrl+C, and checkbox-selected row copy as tab-separated text.
- Ten recent files, clear history, progress and recoverable error messages.

## Architecture and limits

The renderer has no Node.js access. Electron uses sandboxing, context isolation, a narrow typed preload bridge, sender/frame validation, denied permissions, blocked navigation/new windows, and a content security policy. The main process owns the file dialog, clipboard and recent-file storage. No workbook is written, no macro is executed, and no formula is evaluated. Recent paths are stored locally in Electron's user data directory.

SheetJS runs in a worker thread with a 120-second request timeout and a memory limit. Opening reads sheet metadata; selecting a sheet parses that sheet on demand. The worker retains the source bytes and only the latest parsed sheet. Some legacy formats require broader internal parsing in SheetJS. AG Grid virtualizes the active preview; search and sorting operate on its loaded data.

Limits: 200 MB source files; 200,000 rows, 1,024 columns and 1,000,000 cell positions per preview. The UI explicitly reports truncation. Decompressed content may exhaust the worker limit even for small compressed files. Empty sheets show a blank A1. Formatted values and saved formulas are displayed, but cached formula results can be stale or absent. This MVP does not reproduce Excel layout, merged cells, charts, images, conditional formatting, password-protected workbooks, external links, or calculations. CSV encoding and delimiters use SheetJS detection. Column filters operate on displayed text; numeric sorting uses underlying numeric values.

## Verification

`npm test` builds and launches the real Electron app with Playwright. It checks generated fixtures in all five formats, startup opening, sheet switching, formula display, clipboard, search, numeric sorting, recent files, invalid input, renderer isolation and large-sheet truncation. A screenshot is saved to `test-output/viewer.png`.

SheetJS is installed from its [official distribution](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/). The grid uses only [AG Grid Community](https://www.ag-grid.com/react-data-grid/installation/) modules; copying is implemented through Electron rather than an enterprise clipboard module.
