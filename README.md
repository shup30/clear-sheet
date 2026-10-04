# Clear Sheet

Windows spreadsheet editor using Electron, React, TypeScript, Vite, MUI, AG Grid Community, Zustand and SheetJS CE. Opens XLSX, XLS, XLSM, XLSB and CSV.

## Run and verify

```sh
pnpm install --frozen-lockfile
npm run dev
npm run typecheck
npm run lint
npm test
npm run dist
```

`npm start` runs the production build. `release/Clear Sheet Setup 1.0.1.exe` is the unsigned Windows x64 NSIS installer. `lint` runs TypeScript unused-local/parameter checks in both renderer and Electron; no ESLint configuration is installed. Tests build and run workbook/store integrity checks and isolated Electron viewer/editor UI tests. Tests never terminate unrelated Electron processes.

## Application icon

The editable icon source is `public/icon.svg`. `build/icon.png` is its preview and `build/icon.ico` contains the Windows sizes (16, 24, 32, 48, 64, 128 and 256 pixels). The browser preview uses the SVG; the desktop window, executable and installer use the ICO. To use your own Windows icon, replace `build/icon.ico`, run `npm run dist`, close Clear Sheet after saving your work, and run the new installer. Keep `public/icon.svg` and `build/icon.png` in sync when changing the design. Executable resource editing is enabled while code signing remains disabled.

## Editing and saving

Cells and the formula bar edit a sparse workbook model. AG Grid displays values and keeps sorting/filtering separate from worksheet order. Editing, clipboard, formatting, dimensions, visibility, freezing and sheet operations use command history. Undo/redo stores cell deltas and small metadata snapshots; deleting a sheet retains that sheet for undo. The saved history position determines dirty state.

Open commits a new model only after a complete parse. Saving never reloads the file over edits. Sheet switching uses the current model, including empty/cleared sheets. Only a successful save clears dirty state. Mutations are blocked during open/save. Closing unsaved work offers Keep editing or Discard and close; save first to retain changes.

**Excel inputs require an edited copy on their first save.** SheetJS CE cannot preserve every Excel feature; overwriting the source risks losing charts, styles and unsupported metadata. Both renderer and main process protect the original path. Subsequent saves update the edited copy. Reopening that copy starts a new protected session. CSV files can be updated in place.

Save As chooses its destination and format before serialization. Main accepts writes only to authorized destinations, writes an exclusive temporary file beside the destination, flushes it, then renames it. Failure does not truncate the destination. Formulas, supported cached values, comments, hyperlinks, number formats, compatible VBA blobs, dimensions and sheet visibility metadata are retained where SheetJS supports them.

## Important limitations

- XLS/XLSB formula exports are blocked because SheetJS CE drops formulas in those formats. Choose XLSX/XLSM. Macro workbooks cannot be saved into non-macro formats. CSV exports with multiple sheets are blocked. CSV stores plain values, not spreadsheet formulas/formatting.
- Fonts, fills, borders, alignment and freeze panes are session-only with this CE writer. Number formats, row heights, column widths and hidden rows/columns persist. Unsupported Excel content remains in the untouched original; it is not promised in the edited copy.
- The formula engine supports a subset of Excel. Unsupported expressions remain as source formulas and display an error. Original unsupported cached results are retained where possible. No macros, external links or arbitrary code run. Very deep dependency chains are bounded.
- Structural edits involving merged cells, defined names, auto-filter ranges or array formulas are rejected. More complex references such as structured table references are not fully supported.
- The grid materializes the active sheet, capped at 200,000 rows and 1,024 displayed columns. Saving retains the complete model. Full parsing/model construction still happens eagerly in the renderer once per open; very large workbooks can pause the UI. Fully lazy editable parsing needs a larger architectural change.
- Edits and undo/redo patch touched rows and retain column definitions. Structural changes rebuild the active view. History is limited to 200 commands. Range mutations/copy are limited to 200,000 cells, external paste to 500,000 cells. Internal copy preserves raw values and adjusts relative formulas; cut does not implement Excel's dependent-reference retargeting.

## Security

Sandboxed renderer, context isolation, no Node integration, typed preload bridge and validated IPC. No renderer filesystem or generic IPC access. Main-frame checks, denied permissions, blocked navigation/popups and CSP remain enabled. A local `--user-data-dir` argument isolates profiles without exposing filesystem access to the renderer.

## Library references

[SheetJS parsing](https://docs.sheetjs.com/docs/api/parse-options/), [writing](https://docs.sheetjs.com/docs/api/write-options/), [VBA preservation](https://docs.sheetjs.com/docs/csf/features/vba/), and [Electron unload protection](https://www.electronjs.org/docs/latest/api/web-contents#event-will-prevent-unload).


## 1.0.1 interaction fixes

- Save/Save As verifies its preload capability and reports a stale desktop session clearly. Restart the entire Electron process after changing Electron/preload code; Vite updates the UI only. Preserve any unsaved data before restarting an old session. `npm start` now rebuilds both app layers before launching.
- Drag column borders or the bottom edge of a row-number cell to resize. Multi-column Shift-resize is one undo step. Row-height/column-width dialogs apply immediately. Sizes persist through sheet switches and save/reopen using a stable Excel width metric. Zoom changes displayed sizes without modifying stored dimensions.
- Right-click cells or row numbers for row insert/delete and height; right-click column headers for column insert/delete and width. Checkbox-selected disjoint rows delete together, with undo restoring the entire operation. Sorting does not change which source rows are deleted.
- `npm run test:interactions` covers these behaviors in Electron. `npm run test:packaged` verifies both CSV saves and the full interaction workflow against the packaged executable.
