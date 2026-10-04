import React, { useCallback, useEffect, useRef, useState, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { ThemeProvider, CssBaseline, Box, Alert, LinearProgress, Snackbar } from '@mui/material';
import { ModuleRegistry, AllCommunityModule, type GridApi } from 'ag-grid-community';
import type { Row, CellStyle } from '../electron/types';
import { useViewer, type SelRange, type Selection } from './store';
import { editCellsCommand, formatCommand, insertDeleteCommand, sheetOpCommand, inputToChange, compositeCommand } from './workbook/commands';
import { resizeCommand } from './workbook/dimensions';
import { translateFormula } from './workbook/formulas';
import { rangeToTSV, parseTSV } from './workbook/clipboard';
import { DESKTOP_UPDATE_MESSAGE } from './desktop';
import { createAppTheme } from './theme/theme';
import type { ColorMode } from './theme/tokens';
import { TitleBar } from './components/TitleBar';
import { CommandBar } from './components/CommandBar';
import { FormulaBar } from './components/FormulaBar';
import { GridContainer } from './components/GridContainer';
import { SheetTabs } from './components/SheetTabs';
import { StatusBar } from './components/StatusBar';
import { CommandPalette } from './components/CommandPalette';
import { EmptyState } from './components/EmptyState';
import { FormatCellsDialog } from './components/FormatCellsDialog';
import { ChartModal } from './components/ChartModal';
import { FindDialog, RenameDialog, DeleteSheetDialog, DimensionDialog, type FoundItem } from './components/Dialogs';
import { colToLetter, toAddress } from './workbook/cellRef';
import { extractFormulaHighlights } from './workbook/intellisense';
import type { ConditionalFormatRule } from './workbook/conditionalFormatting';
import '@fontsource-variable/inter';
import '@fontsource/jetbrains-mono';
import './style.css';

ModuleRegistry.registerModules([AllCommunityModule]);

function rangeLabel(r: SelRange): string {
  const a = `${colToLetter(r.c1)}${r.r1 + 1}`;
  const b = `${colToLetter(r.c2)}${r.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

function normRange(a: { r: number; c: number }, b: { r: number; c: number }): SelRange {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}

function targetsOf(ranges: SelRange[]): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  if (ranges.reduce((sum, rg) => sum + (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1), 0) > 200000) {
    useViewer.getState().set({ error: 'This operation is limited to 200,000 cells.' });
    return out;
  }
  for (const rg of ranges) {
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    if (area > 200000) continue;
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push({ r, c });
  }
  return out;
}

function replaceLiteral(src: string, find: string, repl: string): string {
  if (!find) return src;
  return src.toLowerCase().split(find.toLowerCase()).join(repl);
}

function App() {
  const s = useViewer();
  const grid = useRef<GridApi<Row> | null>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const anchor = useRef<{ r: number; c: number } | null>(null);
  const shiftDown = useRef(false);
  const localClipboard = useRef<{
    text: string;
    cut: boolean;
    r: number;
    c: number;
    cells: { r: number; c: number; next: { v: string | number | boolean | null; f?: string } }[];
  } | null>(null);

  const selRef = useRef<Selection | null>(s.selection);
  selRef.current = s.selection;

  // Theme state: dark / light with auto-detection
  const [mode, setMode] = useState<ColorMode>(() => {
    const saved = localStorage.getItem('clearsheet_theme');
    if (saved === 'dark' || saved === 'light') return saved;
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  });

  const toggleTheme = useCallback(() => {
    setMode((prev) => {
      const next = prev === 'light' ? 'dark' : 'light';
      localStorage.setItem('clearsheet_theme', next);
      return next;
    });
  }, []);

  const appTheme = useMemo(() => createAppTheme(mode), [mode]);

  const [toast, setToast] = useState('');
  const [drag, setDrag] = useState(false);
  const [query, setQuery] = useState('');
  const [fb, setFb] = useState('');
  const [fbEditing, setFbEditing] = useState(false);
  const [cmdPaletteOpen, setCmdPaletteOpen] = useState(false);

  const displayFb = fbEditing
    ? fb
    : s.selected?.formula !== undefined
    ? '=' + s.selected.formula
    : s.selected?.text ?? '';
  const [findOpen, setFindOpen] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [findScope, setFindScope] = useState<'sheet' | 'book'>('sheet');
  const [matchCase, setMatchCase] = useState(false);
  const [found, setFound] = useState<FoundItem[]>([]);
  const [foundIdx, setFoundIdx] = useState(-1);
  const [renameSheet, setRenameSheet] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState('');
  const [deleteSheet, setDeleteSheet] = useState<string | null>(null);
  const [dimKind, setDimKind] = useState<'colWidth' | 'rowHeight' | null>(null);
  const [dimValue, setDimValue] = useState('');
  const [formatDialogOpen, setFormatDialogOpen] = useState(false);
  const [chartModalOpen, setChartModalOpen] = useState(false);
  const [clipboardRange, setClipboardRange] = useState<{ sheet: string; range: SelRange } | null>(null);
  const [cfRules, setCfRules] = useState<ConditionalFormatRule[]>([]);

  const sheetName = s.sheetName;
  const model = s.model;

  useEffect(() => {
    void s.refresh();
    return window.viewer.onOpen(() => void useViewer.getState().open());
  }, []);

  useEffect(() => {
    const t = setTimeout(() => s.set({ search: query }), 180);
    return () => clearTimeout(t);
  }, [query]);

  useEffect(() => {
    setQuery('');
  }, [s.sheetName, s.book?.name]);

  useEffect(() => {
    const before = (e: BeforeUnloadEvent) => {
      if (useViewer.getState().dirty || useViewer.getState().busy) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', before);
    return () => window.removeEventListener('beforeunload', before);
  }, []);

  const copyText = async (text: string, msg = 'Copied to clipboard') => {
    const result = await window.viewer.copy(text);
    if (result.ok) setToast(msg);
    else useViewer.getState().set({ error: result.error });
    return result.ok;
  };

  const focusCell = useCallback((r: number, c: number) => {
    const api = grid.current;
    if (!api) return;
    const pinned = Array.from({ length: api.getPinnedTopRowCount() }, (_, i) => api.getPinnedTopRow(i)).find(
      (n) => n?.data?.id === r + 1
    );
    const node = pinned ?? api.getRowNode(String(r + 1));
    if (node?.rowIndex == null) return;
    if (!pinned) api.ensureIndexVisible(node.rowIndex);
    api.ensureColumnVisible(String(c));
    api.setFocusedCell(node.rowIndex, String(c), pinned ? 'top' : null);
  }, []);

  const select = useCallback(
    (active: { r: number; c: number }, ranges: SelRange[]) => {
      anchor.current = active;
      useViewer.getState().setSelection({ active, ranges });
      focusCell(active.r, active.c);
    },
    [focusCell]
  );

  const gotoFound = useCallback(
    (f: FoundItem, index: number) => {
      setFoundIdx(index);
      const st = useViewer.getState();
      if (f.sheet !== st.sheetName) {
        void st.load(f.sheet).then(() =>
          select({ r: f.r, c: f.c }, [normRange({ r: f.r, c: f.c }, { r: f.r, c: f.c })])
        );
      } else {
        select({ r: f.r, c: f.c }, [normRange({ r: f.r, c: f.c }, { r: f.r, c: f.c })]);
      }
    },
    [select]
  );

  // ---------- Clipboard operations ----------

  const doCopy = useCallback(
    async (cut: boolean) => {
      const st = useViewer.getState();
      if (!st.model || !st.sheetName || !st.selection) {
        setToast('Select a cell first');
        return;
      }
      const ranges = st.selection.ranges.length
        ? st.selection.ranges
        : [normRange(st.selection.active, st.selection.active)];
      if (ranges.reduce((n, rg) => n + (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1), 0) > 200000) {
        st.set({ error: 'Copy/cut is limited to 200,000 cells.' });
        return;
      }
      const parts = ranges.map((rg) => rangeToTSV(st.model!, st.sheetName!, rg.r1, rg.c1, rg.r2, rg.c2, true));
      const clipboardText = parts.join('\r\n');
      const rg = ranges[0];
      const cells =
        ranges.length === 1
          ? targetsOf(ranges).map(({ r, c }) => {
              const rec = st.model!.getRec(st.sheetName!, r, c);
              return { r, c, next: { v: rec?.v ?? null, f: rec?.f } };
            })
          : [];
      const copied = await copyText(clipboardText, cut ? 'Cut to clipboard' : 'Copied to clipboard');
      if (
        !copied ||
        useViewer.getState().model !== st.model ||
        useViewer.getState().rev !== st.rev ||
        useViewer.getState().busy
      )
        return;
      if (copied && ranges.length > 0 && st.sheetName) {
        setClipboardRange({ sheet: st.sheetName, range: rg });
      }
      localClipboard.current = cells.length ? { text: clipboardText, cut, r: rg.r1, c: rg.c1, cells } : null;
      if (cut) {
        const targets = targetsOf(ranges);
        if (targets.length) {
          st.exec(
            editCellsCommand(
              st.model,
              st.sheetName,
              targets.map((t) => ({ r: t.r, c: t.c, next: { v: null } })),
              'Cut'
            )
          );
        }
      }
    },
    []
  );

  const doPaste = useCallback(async () => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) {
      setToast('Select a cell first');
      return;
    }
    let text = '';
    try {
      const r = await window.viewer.paste();
      if (!r.ok) throw new Error(r.error);
      text = r.data;
    } catch {
      try {
        text = await navigator.clipboard.readText();
      } catch {
        setToast('Clipboard is empty or unavailable');
        return;
      }
    }
    if (
      useViewer.getState().model !== st.model ||
      useViewer.getState().rev !== st.rev ||
      useViewer.getState().busy
    )
      return;
    if (!text) {
      setToast('Clipboard is empty');
      return;
    }
    const gridData = parseTSV(text);
    if (!gridData.length) return;
    const { r: ar, c: ac } = st.selection.active;
    let changes = gridData.flatMap((row, dr) =>
      row.map((t, dc) => ({ r: ar + dr, c: ac + dc, next: inputToChange(t) }))
    );
    const local = localClipboard.current;
    if (local?.text === text) {
      changes = local.cells.map((cell) => ({
        r: ar + cell.r - local.r,
        c: ac + cell.c - local.c,
        next: {
          ...cell.next,
          f:
            cell.next.f === undefined
              ? undefined
              : local.cut
              ? cell.next.f
              : translateFormula(cell.next.f, ar - local.r, ac - local.c),
        },
      }));
    }
    if (changes.length > 500000) {
      st.set({ error: 'Paste is limited to 500,000 cells.' });
      return;
    }
    if (changes.some((c) => c.r >= 1048576 || c.c >= 16384)) {
      st.set({ error: 'Paste exceeds Excel row/column limits.' });
      return;
    }
    st.exec(editCellsCommand(st.model, st.sheetName, changes, 'Paste'));
    setClipboardRange(null);
    const last = changes[changes.length - 1];
    select({ r: last.r, c: last.c }, [{ r1: ar, c1: ac, r2: last.r, c2: last.c }]);
  }, [select]);

  const doDelete = useCallback(() => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const ranges = st.selection.ranges.length
      ? st.selection.ranges
      : [normRange(st.selection.active, st.selection.active)];
    const targets = targetsOf(ranges);
    if (!targets.length) return;
    st.exec(
      editCellsCommand(
        st.model,
        st.sheetName,
        targets.map((t) => ({ r: t.r, c: t.c, next: { v: null } })),
        'Clear cells'
      )
    );
  }, []);

  const handleAutoSum = useCallback(() => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) return;
    const { r, c } = st.selection.active;
    let rTop = r - 1;
    while (rTop >= 0) {
      const disp = st.model.displayOf(st.sheetName, rTop, c);
      if (typeof disp.v === 'number' || disp.formula) rTop--;
      else break;
    }
    rTop++;
    if (rTop < r) {
      setFb(`=SUM(${toAddress(rTop, c)}:${toAddress(r - 1, c)})`);
      setFbEditing(true);
      return;
    }
    let cLeft = c - 1;
    while (cLeft >= 0) {
      const disp = st.model.displayOf(st.sheetName, r, cLeft);
      if (typeof disp.v === 'number' || disp.formula) cLeft--;
      else break;
    }
    cLeft++;
    if (cLeft < c) {
      setFb(`=SUM(${toAddress(r, cLeft)}:${toAddress(r, c - 1)})`);
      setFbEditing(true);
      return;
    }
    setFb('=SUM()');
    setFbEditing(true);
  }, []);

  const insertDirectValue = useCallback((val: string) => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) return;
    const { r, c } = st.selection.active;
    st.exec(
      editCellsCommand(
        st.model,
        st.sheetName,
        [{ r, c, next: inputToChange(val) }],
        'Insert value'
      )
    );
  }, []);

  const switchSheetDelta = useCallback((delta: number) => {
    const st = useViewer.getState();
    if (!st.book || !st.sheetName) return;
    const sheets = st.book.sheets;
    const curIdx = sheets.indexOf(st.sheetName);
    if (curIdx < 0) return;
    const nextIdx = (curIdx + delta + sheets.length) % sheets.length;
    void st.load(sheets[nextIdx]);
  }, []);

  const handleFormulaRangeSelect = useCallback((rangeStr: string) => {
    setFb((prev) => {
      if (prev === '' || prev === '=') return `=${rangeStr}`;
      if (/[\(\+\-\*\/\^\&\,=]\s*$/.test(prev)) {
        return `${prev}${rangeStr}`;
      }
      const updated = prev.replace(/(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?\s*$/, rangeStr);
      if (updated !== prev) return updated;
      return `${prev}, ${rangeStr}`;
    });
  }, []);

  const handleAddCf = useCallback((type: 'gt' | 'colorScale' | 'dataBar') => {
    const st = useViewer.getState();
    if (!st.sheetName || !st.selection) return;
    const rg = st.selection.ranges[0] ?? normRange(st.selection.active, st.selection.active);
    const id = `cf_${Date.now()}`;
    if (type === 'colorScale') {
      setCfRules((prev) => [
        ...prev,
        {
          id,
          sheet: st.sheetName!,
          range: rg,
          ruleType: 'colorScale',
          minColor: '#f87171',
          midColor: '#fef08a',
          maxColor: '#4ade80',
        },
      ]);
      setToast('Applied Color Scale Heatmap');
    } else if (type === 'dataBar') {
      setCfRules((prev) => [
        ...prev,
        {
          id,
          sheet: st.sheetName!,
          range: rg,
          ruleType: 'dataBar',
          barColor: '#38bdf8',
        },
      ]);
      setToast('Applied Data Bars');
    } else if (type === 'gt') {
      setCfRules((prev) => [
        ...prev,
        {
          id,
          sheet: st.sheetName!,
          range: rg,
          ruleType: 'highlight',
          operator: 'gt',
          val1: 0,
          bg: '#dcfce7',
          color: '#166534',
        },
      ]);
      setToast('Applied Highlight > 0');
    }
  }, []);

  const handleClearCf = useCallback(() => {
    setCfRules([]);
    setToast('Cleared Conditional Formatting');
  }, []);

  // Global keyboard shortcuts
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const st = useViewer.getState();
      shiftDown.current = e.shiftKey;
      const target = e.target as HTMLElement;
      const inField =
        target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target.isContentEditable;
      const gridEditing = !!document.querySelector('.ag-cell-inline-editing');
      if (e.key === 'Shift') shiftDown.current = true;
      if (st.busy) return;

      // Command palette: Ctrl+K
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setCmdPaletteOpen((prev) => !prev);
        return;
      }

      if (target.closest('[role=dialog]')) return;

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f' && !e.shiftKey) {
        e.preventDefault();
        searchRef.current?.focus();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'h') {
        e.preventDefault();
        setFindOpen(true);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (!st.model) return;
        grid.current?.stopEditing();
        if (document.activeElement instanceof HTMLInputElement) document.activeElement.blur();
        if (e.shiftKey) void useViewer.getState().saveAs();
        else void useViewer.getState().save();
        return;
      }
      if (inField || gridEditing) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
        e.preventDefault();
        st.undo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) {
        e.preventDefault();
        st.redo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        if (st.selection) {
          e.preventDefault();
          void doCopy(false);
        }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
        if (st.selection) {
          e.preventDefault();
          void doCopy(true);
        }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
        if (st.selection && st.model) {
          e.preventDefault();
          void doPaste();
        }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
        if (st.model && st.sheetName) {
          e.preventDefault();
          const sh = st.model.getSheet(st.sheetName);
          if (sh)
            select(
              { r: 0, c: 0 },
              [{ r1: 0, c1: 0, r2: Math.max(0, sh.rowCount - 1), c2: Math.max(0, sh.colCount - 1) }]
            );
        }
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === '1') {
        e.preventDefault();
        setFormatDialogOpen(true);
        return;
      }
      if (e.altKey && (e.key === '=' || e.key === '+')) {
        e.preventDefault();
        handleAutoSum();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key === ';') {
        e.preventDefault();
        const now = new Date();
        const ds = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
        insertDirectValue(ds);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === ':' || (e.shiftKey && e.key === ';'))) {
        e.preventDefault();
        const now = new Date();
        const ts = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
        insertDirectValue(ts);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'PageDown') {
        e.preventDefault();
        switchSheetDelta(1);
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'PageUp') {
        e.preventDefault();
        switchSheetDelta(-1);
        return;
      }
      if (e.key === 'Escape') {
        setClipboardRange(null);
      }
      if ((e.key === 'Delete' || e.key === 'Backspace') && st.selection && st.model) {
        e.preventDefault();
        doDelete();
      }
    };
    const up = (e: KeyboardEvent) => {
      if (e.key === 'Shift') shiftDown.current = false;
    };
    window.addEventListener('keydown', handler);
    window.addEventListener('keyup', up);
    return () => {
      window.removeEventListener('keydown', handler);
      window.removeEventListener('keyup', up);
    };
  }, [doCopy, doPaste, doDelete, select]);

  const commitFormulaBar = () => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const { r, c } = st.selection.active;
    const cur = st.model.displayOf(st.sheetName, r, c);
    const curEdit = cur.formula !== undefined ? '=' + cur.formula : cur.text;
    if (fb === curEdit) {
      setFbEditing(false);
      return;
    }
    try {
      st.exec(
        editCellsCommand(
          st.model,
          st.sheetName,
          [{ r, c, next: inputToChange(fb, st.model.getRec(st.sheetName, r, c)?.style?.numFmt === '@') }],
          'Edit cell'
        )
      );
    } catch (e) {
      st.set({ error: String(e) });
    }
    setFbEditing(false);
  };

  const fmt = (patch: Partial<CellStyle> | null, numFmt?: string, label = 'Format') => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) {
      setToast('Select cells first');
      return;
    }
    const ranges = st.selection.ranges.length
      ? st.selection.ranges
      : [normRange(st.selection.active, st.selection.active)];
    const targets = targetsOf(ranges);
    if (!targets.length) return;
    st.exec(formatCommand(st.model, st.sheetName, targets, patch ?? null, numFmt, label));
  };

  const insertDel = (axis: 'row' | 'col', mode: 'insert' | 'delete') => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) {
      setToast('Select a cell first');
      return;
    }
    const at =
      axis === 'row'
        ? Math.min(...st.selection.ranges.flatMap((rg) => [rg.r1]))
        : Math.min(...st.selection.ranges.flatMap((rg) => [rg.c1]));
    const span =
      axis === 'row'
        ? Math.max(...st.selection.ranges.map((rg) => rg.r2 - rg.r1 + 1))
        : Math.max(...st.selection.ranges.map((rg) => rg.c2 - rg.c1 + 1));
    const count = st.selection.ranges.length > 1 || span > 1 ? span : 1;
    try {
      st.exec(insertDeleteCommand(st.model, st.sheetName, axis, mode, at, count));
    } catch (e) {
      st.set({ error: e instanceof Error ? e.message : String(e) });
    }
    setToast(`${mode === 'insert' ? 'Inserted' : 'Deleted'} ${count} ${axis}(s)`);
  };

  const hideUnhide = (axis: 'row' | 'col', hide: boolean) => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const sh = st.model.getSheet(st.sheetName)!;
    const ranges = st.selection.ranges.length
      ? st.selection.ranges
      : [normRange(st.selection.active, st.selection.active)];
    sh.hiddenRows = new Set(sh.hiddenRows);
    sh.hiddenCols = new Set(sh.hiddenCols);
    for (const rg of ranges) {
      if (axis === 'row')
        for (let r = rg.r1; r <= rg.r2; r++) {
          if (hide) sh.hiddenRows.add(r);
          else sh.hiddenRows.delete(r);
        }
      else
        for (let c = rg.c1; c <= rg.c2; c++) {
          if (hide) sh.hiddenCols.add(c);
          else sh.hiddenCols.delete(c);
        }
    }
    st.set({ layoutRev: st.layoutRev + 1 });
  };

  const unhideAll = () => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const sh = st.model.getSheet(st.sheetName)!;
    sh.hiddenRows.clear();
    sh.hiddenCols.clear();
    st.set({ layoutRev: st.layoutRev + 1 });
  };

  const freeze = (rows: number, cols: number) => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const sh = st.model.getSheet(st.sheetName)!;
    sh.frozenRows = rows;
    sh.frozenCols = cols;
    st.set({ layoutRev: st.layoutRev + 1 });
    setToast(rows || cols ? `Frozen (${rows}R × ${cols}C)` : 'Freeze cleared');
  };

  const freezeAtSelection = () => {
    const st = useViewer.getState();
    if (!st.selection) return;
    freeze(st.selection.active.r, st.selection.active.c);
  };

  const copyRows = () => {
    const colCount = model && sheetName ? Math.min(Math.max(model.getSheet(sheetName)?.colCount ?? 10, 10), 1024) : 0;
    const rows = grid.current?.getSelectedNodes().sort((a, b) => (a.rowIndex ?? 0) - (b.rowIndex ?? 0)) ?? [];
    if (!rows.length) {
      setToast('Select rows using the checkboxes first');
      return;
    }
    const text = rows
      .map((n) =>
        Array.from({ length: colCount }, (_, i) => {
          const t = (n.data as Row | undefined)?.cells[String(i)]?.text ?? '';
          return /[\t\r\n"]/.test(t) ? '"' + t.replaceAll('"', '""') + '"' : t;
        }).join('\t')
      )
      .join('\r\n');
    void copyText(text);
  };

  // Find & Replace logic
  const runFind = (all: boolean): FoundItem[] => {
    const st = useViewer.getState();
    if (!st.model || !findText) return [];
    const q = matchCase ? findText : findText.toLowerCase();
    const sheets = findScope === 'sheet' ? [st.sheetName!] : st.model.sheetNames();
    const out: FoundItem[] = [];
    for (const name of sheets) {
      const sh = st.model.getSheet(name);
      if (!sh) continue;
      for (const [k, rec] of sh.cells) {
        const r = Math.floor(k / 16384),
          c = k % 16384;
        const d = st.model.displayOf(name, r, c);
        const hay = [d.text, rec.f !== undefined ? '=' + rec.f : '', typeof rec.v === 'string' ? rec.v : ''];
        if (hay.some((h) => (matchCase ? h : h.toLowerCase()).includes(q))) {
          out.push({ sheet: name, r, c, text: d.text });
          if (!all && out.length >= 5000) break;
        }
      }
    }
    out.sort((a, b) => (a.sheet === b.sheet ? a.r - b.r || a.c - b.c : a.sheet.localeCompare(b.sheet)));
    return out;
  };

  const findNav = (dir: 1 | -1) => {
    let list = found;
    if (!list.length) {
      list = runFind(false);
      setFound(list);
    }
    if (!list.length) {
      setToast('No matches');
      return;
    }
    let i = foundIdx;
    i = i < 0 ? (dir === 1 ? 0 : list.length - 1) : (i + dir + list.length) % list.length;
    setFoundIdx(i);
    gotoFound(list[i], i);
  };

  const doReplace = (all: boolean) => {
    const st = useViewer.getState();
    if (!st.model || !findText) return;
    const list = all
      ? runFind(true)
      : found.length
      ? [found[Math.max(0, foundIdx)]]
      : runFind(false).slice(0, 1);
    if (!list.length) {
      setToast('No matches');
      return;
    }
    const q = matchCase ? findText : findText.toLowerCase();
    const bySheet = new Map<
      string,
      { r: number; c: number; next: { v: string | number | boolean | null; f?: string } }[]
    >();
    for (const f of list) {
      const rec = st.model.getRec(f.sheet, f.r, f.c);
      if (!rec) continue;
      let nv: string;
      if (rec.f !== undefined) {
        const src = '=' + rec.f;
        nv = matchCase ? src.replaceAll(findText, replaceText) : replaceLiteral(src, findText, replaceText);
        if (nv === src && !all) continue;
      } else if (typeof rec.v === 'string') {
        nv = matchCase ? rec.v.replaceAll(findText, replaceText) : replaceLiteral(rec.v, findText, replaceText);
        if (nv === rec.v && !all) continue;
      } else {
        const t = st.model.displayOf(f.sheet, f.r, f.c).text;
        if (!(matchCase ? t : t.toLowerCase()).includes(q)) continue;
        const raw = String(rec.v ?? '');
        nv = matchCase ? raw.replaceAll(findText, replaceText) : replaceLiteral(raw, findText, replaceText);
        if (nv === raw) continue;
      }
      if (!bySheet.has(f.sheet)) bySheet.set(f.sheet, []);
      bySheet
        .get(f.sheet)!
        .push({
          r: f.r,
          c: f.c,
          next:
            rec.f === undefined && typeof rec.v === 'string'
              ? { v: nv }
              : inputToChange(nv),
        });
    }
    let n = 0;
    const commands = [];
    for (const [name, changes] of bySheet) {
      if (!changes.length) continue;
      commands.push(editCellsCommand(st.model, name, changes, all ? 'Replace all' : 'Replace'));
      n += changes.length;
    }
    if (commands.length) st.exec(compositeCommand(commands, all ? 'Replace all' : 'Replace'));
    setToast(n ? `Replaced ${n} cell(s)` : 'No matches');
    if (!all && n) {
      const l = runFind(false);
      setFound(l);
      setFoundIdx(-1);
    }
  };

  const sheetOp = (op: Parameters<typeof sheetOpCommand>[1]) => {
    const st = useViewer.getState();
    if (!st.model) return;
    if (op.type === 'delete') {
      setDeleteSheet(op.name);
      return;
    }
    if (op.type === 'rename') {
      setRenameSheet(op.oldName);
      setRenameValue(op.oldName);
      return;
    }
    const before = op.type === 'add' || op.type === 'duplicate' ? st.model.sheetNames() : null;
    try {
      st.exec(sheetOpCommand(st.model, op));
    } catch (e) {
      st.set({ error: e instanceof Error ? e.message : String(e) });
      return;
    }
    if (before) {
      const created = useViewer.getState().model?.sheetNames().find((n) => !before.includes(n));
      if (created) void useViewer.getState().load(created);
    }
  };

  const confirmRename = () => {
    const st = useViewer.getState();
    if (!st.model || !renameSheet) {
      setRenameSheet(null);
      return;
    }
    const nn = renameValue.trim();
    if (nn && nn !== renameSheet) {
      try {
        st.exec(sheetOpCommand(st.model, { type: 'rename', oldName: renameSheet, newName: nn }));
      } catch (e) {
        st.set({ error: String(e) });
        return;
      }
    }
    setRenameSheet(null);
  };

  const confirmDeleteSheet = () => {
    const st = useViewer.getState();
    if (st.model && deleteSheet) {
      try {
        st.exec(sheetOpCommand(st.model, { type: 'delete', name: deleteSheet }));
      } catch (e) {
        st.set({ error: String(e) });
      }
    }
    setDeleteSheet(null);
  };

  const applyDim = () => {
    const st = useViewer.getState();
    if (st.model && st.sheetName && st.selection && dimKind) {
      const selection = st.selection;
      const sizes = new Map<number, number>();
      const axis = dimKind === 'colWidth' ? 'col' : 'row';
      const n =
        axis === 'col'
          ? Math.max(5, Math.min(200, Number(dimValue) || 12))
          : Math.max(20, Math.min(600, Number(dimValue) || 30));
      for (const range of selection.ranges) {
        for (let i = axis === 'col' ? range.c1 : range.r1; i <= (axis === 'col' ? range.c2 : range.r2); i++) {
          sizes.set(i, n);
        }
      }
      st.exec(resizeCommand(st.model, st.sheetName, axis, sizes));
    }
    setDimKind(null);
  };

  const selLabel = s.selection
    ? (s.selection.ranges.length > 1
        ? `${s.selection.ranges.length} ranges`
        : rangeLabel(s.selection.ranges[0] ?? normRange(s.selection.active, s.selection.active)))
    : (s.selected?.address ?? '—');

  const activeStyle =
    s.selection && s.model && s.sheetName
      ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style
      : undefined;

  return (
    <ThemeProvider theme={appTheme}>
      <CssBaseline />
      <Box
        className="app"
        onDragOver={(e) => {
          e.preventDefault();
          setDrag(true);
        }}
        onDragLeave={(e) => {
          if (!e.currentTarget.contains(e.relatedTarget as Node)) setDrag(false);
        }}
        onDrop={(e) => {
          e.preventDefault();
          setDrag(false);
          if (e.dataTransfer.files[0]) void s.open(undefined, e.dataTransfer.files[0]);
        }}
      >
        {typeof window.viewer.chooseSavePath !== 'function' && (
          <Alert severity="error">{DESKTOP_UPDATE_MESSAGE}</Alert>
        )}

        {/* 1. Windows 11 Native TitleBar */}
        <TitleBar
          bookName={s.book?.name ?? null}
          filePath={s.filePath}
          dirty={s.dirty}
          busy={s.busy}
          mode={mode}
          onToggleTheme={toggleTheme}
          onOpen={() => void s.open()}
          onSave={() => {
            grid.current?.stopEditing();
            void useViewer.getState().save();
          }}
          onSaveAs={() => {
            grid.current?.stopEditing();
            void useViewer.getState().saveAs();
          }}
          onOpenCommandPalette={() => setCmdPaletteOpen(true)}
        />

        {s.error && (
          <Alert severity="error" onClose={() => s.set({ error: '' })}>
            {s.error}
          </Alert>
        )}

        {s.busy && <LinearProgress color="primary" sx={{ height: 2 }} />}

        {s.book ? (
          <>
            {/* 2. Grouped CommandBar (Ribbon) */}
            <CommandBar
              canUndo={s.canUndo}
              canRedo={s.canRedo}
              hasSelection={!!s.selection}
              hasSelectedCell={!!s.selected}
              busy={s.busy}
              zoom={s.zoom}
              query={query}
              searchRef={searchRef}
              onUndo={s.undo}
              onRedo={s.redo}
              onCopy={doCopy}
              onPaste={doPaste}
              onCopyCell={() => void copyText(s.selected?.text ?? '')}
              onCopyRows={copyRows}
              onFormat={fmt}
              onInsertDel={insertDel}
              onOpenDim={(kind) => {
                setDimKind(kind);
                setDimValue(kind === 'colWidth' ? '12' : '30');
              }}
              onHideUnhide={hideUnhide}
              onUnhideAll={unhideAll}
              onResetFilters={() => {
                grid.current?.setFilterModel(null);
                grid.current?.applyColumnState({ defaultState: { sort: null } });
                setQuery('');
              }}
              onFreeze={freeze}
              onFreezeAtSelection={freezeAtSelection}
              onOpenFind={() => setFindOpen(true)}
              onSearchChange={setQuery}
              onZoomChange={(zoom) => s.set({ zoom })}
              isBoldActive={!!activeStyle?.bold}
              isItalicActive={!!activeStyle?.italic}
              isUnderlineActive={!!activeStyle?.underline}
              isWrapActive={!!activeStyle?.wrap}
              onOpenFormatCells={() => setFormatDialogOpen(true)}
              onOpenChart={() => setChartModalOpen(true)}
              onAutoSum={handleAutoSum}
            />

            {/* 3. JetBrains Mono FormulaBar */}
            <FormulaBar
              address={selLabel}
              value={displayFb}
              isEditing={fbEditing}
              onFocus={() => {
                setFb(displayFb);
                setFbEditing(true);
              }}
              onBlur={() => {
                setFbEditing(false);
              }}
              onChange={(val) => {
                setFb(val);
                setFbEditing(true);
              }}
              onCommit={() => {
                commitFormulaBar();
                if (s.selection) focusCell(s.selection.active.r, s.selection.active.c);
              }}
              onCancel={() => {
                setFbEditing(false);
              }}
            />

            {s.model?.requiresCopy && (
              <Alert severity="info" sx={{ py: 0.5, fontSize: '0.75rem' }}>
                Save creates an edited copy to preserve the original Excel file. Fonts, fills, borders, freeze panes,
                charts and other advanced features may not persist in that copy.
              </Alert>
            )}

            {/* 4. AG Grid Quartz Viewport */}
            <GridContainer
              mode={mode}
              gridRef={grid}
              selRef={selRef}
              anchorRef={anchor}
              shiftDownRef={shiftDown}
              onSelect={select}
              onFocusCell={focusCell}
              onOpenDimWithValues={(kind, val) => {
                setDimKind(kind);
                setDimValue(val);
              }}
              isFormulaEditing={fbEditing && fb.startsWith('=')}
              formulaHighlights={fbEditing && fb.startsWith('=') ? extractFormulaHighlights(fb) : []}
              onFormulaRangeSelect={handleFormulaRangeSelect}
              clipboardRange={clipboardRange}
              cfRules={cfRules}
            />

            {/* 5. Bottom Sheet Tabs */}
            <SheetTabs
              sheetNames={s.book.sheets}
              activeSheet={sheetName}
              busy={s.busy}
              onSelectSheet={(name) => void s.load(name)}
              onAddSheet={() => sheetOp({ type: 'add' })}
              onRenameSheet={(name) => {
                setRenameSheet(name);
                setRenameValue(name);
              }}
              onDuplicateSheet={(name) => sheetOp({ type: 'duplicate', name })}
              onMoveSheet={(name, dir) => {
                const st = useViewer.getState();
                if (st.model) {
                  const i = st.model.sheetNames().indexOf(name);
                  const to = dir === 'left' ? Math.max(0, i - 1) : i + 1;
                  st.exec(sheetOpCommand(st.model, { type: 'move', name, from: i, to }));
                }
              }}
              onDeleteSheet={(name) => sheetOp({ type: 'delete', name })}
            />

            {/* 6. Status Bar with live metrics */}
            <StatusBar
              sheetName={sheetName}
              model={model}
              dirty={s.dirty}
              selLabel={selLabel}
              selection={s.selection}
            />
          </>
        ) : (
          <EmptyState
            busy={s.busy}
            recents={s.recents}
            onOpen={(p) => void s.open(p)}
            onClearRecent={async () => {
              const r = await window.viewer.clearRecent();
              if (r.ok) void s.refresh();
              else s.set({ error: r.error });
            }}
          />
        )}

        {/* Command Palette (Ctrl+K) */}
        <CommandPalette
          open={cmdPaletteOpen}
          onClose={() => setCmdPaletteOpen(false)}
          mode={mode}
          onToggleTheme={toggleTheme}
          onSave={() => {
            grid.current?.stopEditing();
            void useViewer.getState().save();
          }}
          onSaveAs={() => {
            grid.current?.stopEditing();
            void useViewer.getState().saveAs();
          }}
          onOpen={() => void s.open()}
          onOpenFind={() => setFindOpen(true)}
          onUndo={s.undo}
          onRedo={s.redo}
          onAddSheet={() => sheetOp({ type: 'add' })}
          onFreezeFirstRow={() => freeze(1, 0)}
          onFreezeFirstCol={() => freeze(0, 1)}
          onUnfreeze={() => freeze(0, 0)}
          onResetFilters={() => {
            grid.current?.setFilterModel(null);
            grid.current?.applyColumnState({ defaultState: { sort: null } });
            setQuery('');
          }}
          onClearFormatting={() => fmt(null, undefined, 'Clear formatting')}
          onSetZoom={(zoom) => s.set({ zoom })}
          onFormatCells={() => setFormatDialogOpen(true)}
          onInsertChart={() => setChartModalOpen(true)}
          onAutoSum={handleAutoSum}
          onAddCf={handleAddCf}
          onClearCf={handleClearCf}
        />

        {/* Modal Dialogs */}
        <FindDialog
          open={findOpen}
          onClose={() => setFindOpen(false)}
          findText={findText}
          onFindTextChange={(val) => {
            setFindText(val);
            setFound([]);
            setFoundIdx(-1);
          }}
          replaceText={replaceText}
          onReplaceTextChange={setReplaceText}
          findScope={findScope}
          onFindScopeChange={setFindScope}
          matchCase={matchCase}
          onMatchCaseChange={setMatchCase}
          found={found}
          foundIdx={foundIdx}
          onFindNav={findNav}
          onFindAll={() => {
            const l = runFind(true);
            setFound(l);
            setFoundIdx(l.length ? 0 : -1);
            if (!l.length) setToast('No matches');
          }}
          onReplace={doReplace}
          onGotoFound={gotoFound}
        />

        <RenameDialog
          open={renameSheet !== null}
          value={renameValue}
          onChange={setRenameValue}
          onConfirm={confirmRename}
          onCancel={() => setRenameSheet(null)}
        />

        <DeleteSheetDialog
          open={deleteSheet !== null}
          sheetName={deleteSheet}
          onConfirm={confirmDeleteSheet}
          onCancel={() => setDeleteSheet(null)}
        />

        <DimensionDialog
          kind={dimKind}
          value={dimValue}
          onChange={setDimValue}
          onConfirm={applyDim}
          onCancel={() => setDimKind(null)}
        />

        {/* Master Format Cells Dialog (Ctrl+1) */}
        <FormatCellsDialog
          open={formatDialogOpen}
          activeValue={s.selected?.text}
          currentStyle={activeStyle}
          currentNumFmt={s.selection && s.model && s.sheetName ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style?.numFmt : undefined}
          onClose={() => setFormatDialogOpen(false)}
          onApply={(patch, numFmt) => fmt(patch, numFmt, 'Format cells')}
        />

        {/* Native Interactive SVG Chart Modal */}
        <ChartModal
          open={chartModalOpen}
          onClose={() => setChartModalOpen(false)}
          sheetName={sheetName}
          selection={s.selection}
          model={model}
        />

        {drag && <div className="drop-overlay">Drop spreadsheet to open</div>}
        <Snackbar open={!!toast} autoHideDuration={2200} onClose={() => setToast('')} message={toast} />
      </Box>
    </ThemeProvider>
  );
}

createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
