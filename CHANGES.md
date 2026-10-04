# Excel Viewer → Editor: Change Reference

## What it is now
Lightweight Excel **editor** (was read-only viewer). Electron + React + TS + Vite + MUI + AG Grid Community + Zustand + SheetJS. Single-window Windows app (`Clear Sheet`).

## New architecture
- AG Grid is view-only. Source of truth is `WorkbookModel` (`src/workbook/model.ts`, sparse `Map` cells, per-sheet dims/widths/hidden/freeze).
- Original SheetJS `WorkBook` is kept inside the model; save syncs deltas back into it (preserves charts/macros/etc. as far as SheetJS supports).
- Edits flow: grid `valueSetter` / formula bar / toolbar → `Command` → `store.exec` → incremental grid patch via `cmd.touched` (full rebuild only for structural ops).
- Undo/redo is delta-based (`src/workbook/commands.ts`), never clones the workbook.
- Custom formula engine (`src/workbook/formulas.ts`, no new deps): SUM/AVERAGE/MIN/MAX/COUNT(A)/IF/AND/OR/ROUND/CONCAT/LEFT/RIGHT/MID/LEN/VLOOKUP/XLOOKUP(exact)/INDEX/MATCH/date fns, cross-sheet refs, dependency-tracked incremental recalc, `#DIV/0!/#REF!/#CYCLE!` errors.

## Files
- Added: `src/workbook/{cellRef,format,formulas,model,commands,clipboard}.ts`, `tests/editor.cjs`.
- Rewrote: `src/main.tsx` (ribbon, formula bar, grid, tabs, find/replace, dialogs), `src/store.ts` (editor store, dirty, save/saveAs, undo/redo), `src/style.css` (+ribbon).
- Extended: `electron/{main,preload,types}.ts` — new typed IPC `viewer:bytes|save|saveAs|paste|getPath`. Atomic save (tmp+rename). Security unchanged (isolation, sandbox, no renderer fs).
- Fixed: `electron/worker.ts` sizes previews from `!fullref` (sparse sheets); `tests/smoke.cjs` one flaky timing assert made deterministic.

## Key behaviors / gotchas
- Dirty flag per model; `beforeunload` guard + `confirmDiscard` on open. **Tests must save before `app.close()`** — CDP cannot dismiss Electron's native beforeunload dialog, close hangs.
- `window.prompt` is unsupported in this Electron setup — use MUI dialogs (rename/delete sheet, row height/col width already use them). `window.confirm` works.
- SheetJS CE drops font/fill/border/alignment on write; number formats (`cell.z`) persist. In-session formatting is full-fidelity.
- Grid `valueGetter` returns edit text (`=formula` or raw), `valueFormatter` shows display text; comparator is numeric-aware (keeps smoke numeric-sort test green).
- Keep these stable for smoke: `Copy cell` button, `Search sheet` label, `Reset filters`, `Formula bar` label, `.address` element, tab names = sheet names.
- Test hygiene: kill stale `electron.exe` before runs (orphans hold single-instance lock); Playwright selectors must be unique (chip text appears in header AND footer); after clicking a cell, wait for `.address` to match before reading dependent UI.

## Commands
- `npm run build` — typecheck + vite + electron build.
- `node tests/smoke.cjs` — viewer regression (needs build first).
- `node tests/editor.cjs` — editor E2E (edit/recalc/undo/clipboard/format/sheets/freeze/find/save/reopen).
- `npm run dist` — Windows x64 NSIS installer → `release/Clear Sheet Setup 1.0.0.exe`.

## Senior review fixes (2026-09-18)

- Replaced preview/save reload races with transactional full-model open and model-only sheet switching. Saves serialize the same edited model; edits are locked while saving.
- Save As chooses format first. Writes validate authorized paths, flush an exclusive temporary file, then rename. First Excel save requires a separate copy to preserve unsupported source content. XLS/XLSB formula exports and multi-sheet CSV exports are rejected rather than losing data.
- Serialization keeps cell metadata, valid numeric Excel error codes, number formats, dimension metadata, worksheet visibility and VBA where supported. It never mutates the original workbook while serializing.
- Structural undo restores deleted cells, formulas (including cross-sheet/absolute references), dimensions and hidden state. Rename preserves string literals. History tracks its saved position and supports layout operations. Grid updates are incremental for undo/redo as well as edits.
- Fixed formula dependency ordering, cycles, parse errors becoming zero, case-insensitive dependencies, absolute references, sparse lookup positions, and formula-copy offsets. Percent/date input hints now reach the model.
- Fixed clipboard failure/stale-operation guards, raw internal clipboard types, frozen-row duplication/focus, formula sorting and editor shortcuts. Added native unsaved-close confirmation.
- Added integrity regressions, Save As/close E2E cases and packaged CSV smoke test. Tests isolate profiles and do not kill unrelated Electron processes.
- Remaining compatibility/performance limits are documented in README.md; CE visual styles/freeze are explicitly session-only. Full editable parsing remains eager.
