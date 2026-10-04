import React, { useMemo, useCallback, useEffect, useRef, useState } from 'react';
import { AgGridReact } from 'ag-grid-react';
import type {
  ColDef, GridApi, CellFocusedEvent, CellClickedEvent,
  CellClassParams, RowHeightParams, ColumnResizedEvent, CellContextMenuEvent
} from 'ag-grid-community';
import { Menu, MenuItem, Divider } from '@mui/material';
import type { Row } from '../../electron/types';
import { useViewer, type SelRange, type Selection } from '../store';
import { colToLetter, toAddress } from '../workbook/cellRef';
import { RowHeader } from '../RowHeader';
import { resizeCommand } from '../workbook/dimensions';
import { axisCommand, mergeSpans, type Span } from '../workbook/axisOperations';
import { inputToChange, editCellsCommand } from '../workbook/commands';
import { getGridTheme, type ColorMode, palette } from '../theme/tokens';
import { translateFormula } from '../workbook/formulas';
import { evaluateConditionalFormatting, type ConditionalFormatRule } from '../workbook/conditionalFormatting';

interface GridContainerProps {
  mode: ColorMode;
  gridRef: React.MutableRefObject<GridApi<Row> | null>;
  selRef: React.MutableRefObject<Selection | null>;
  anchorRef: React.MutableRefObject<{ r: number; c: number } | null>;
  shiftDownRef: React.MutableRefObject<boolean>;
  onSelect: (active: { r: number; c: number }, ranges: SelRange[]) => void;
  onFocusCell: (r: number, c: number) => void;
  onOpenDimWithValues: (kind: 'colWidth' | 'rowHeight', val: string) => void;
  isFormulaEditing?: boolean;
  formulaHighlights?: { range: { r1: number; c1: number; r2: number; c2: number }; color: string }[];
  onFormulaRangeSelect?: (rangeStr: string) => void;
  clipboardRange?: { sheet: string; range: SelRange } | null;
  cfRules?: ConditionalFormatRule[];
}

function normRange(a: { r: number; c: number }, b: { r: number; c: number }): SelRange {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}

export const GridContainer: React.FC<GridContainerProps> = ({
  mode,
  gridRef,
  selRef,
  anchorRef,
  shiftDownRef,
  onSelect,
  onOpenDimWithValues,
  isFormulaEditing = false,
  formulaHighlights = [],
  onFormulaRangeSelect,
  clipboardRange = null,
  cfRules = [],
}) => {
  const s = useViewer();
  const containerRef = useRef<HTMLDivElement>(null);
  const [context, setContext] = useState<{ x: number; y: number; rows: Span[]; cols: Span[] } | null>(null);

  // Fill handle position and drag state
  const [handlePos, setHandlePos] = useState<{
    top: number;
    left: number;
    r1: number;
    c1: number;
    r2: number;
    c2: number;
  } | null>(null);
  const [dragPreview, setDragPreview] = useState<{ top: number; left: number; width: number; height: number } | null>(null);

  const model = s.model;
  const sheetName = s.sheetName;
  const colCount = model && sheetName ? Math.min(Math.max(model.getSheet(sheetName)?.colCount ?? 10, 10), 1024) : 0;
  const frozenCols = model && sheetName ? model.getSheet(sheetName)?.frozenCols ?? 0 : 0;

  const currentPalette = palette[mode];

  // Helper for computing numeric stats for CF color scales & data bars
  const getRangeStats = useCallback((sh: string, rg: { r1: number; c1: number; r2: number; c2: number }) => {
    if (!model) return { min: 0, max: 10, mid: 5 };
    const nums: number[] = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const v = model.displayOf(sh, r, c).v;
        if (typeof v === 'number') nums.push(v);
      }
    }
    if (!nums.length) return { min: 0, max: 10, mid: 5 };
    const min = Math.min(...nums);
    const max = Math.max(...nums);
    return { min, max, mid: (min + max) / 2 };
  }, [model]);

  // Cell style logic: user styles + selection state + numeric alignment + error styling + CF + formula highlights
  const cellStyleFn = useCallback((p: CellClassParams<Row>) => {
    const out: Record<string, string> = {
      fontWeight: 'normal',
      fontStyle: 'normal',
      textDecoration: 'none',
      fontFamily: 'inherit',
      fontSize: 'inherit',
      color: 'inherit',
      backgroundColor: '',
      textAlign: 'left',
      verticalAlign: 'middle',
      whiteSpace: 'nowrap',
      lineHeight: 'normal',
      borderTop: '',
      borderBottom: '',
      borderLeft: '',
      borderRight: '',
      outline: '',
      boxShadow: '',
    };

    const cellData = p.data?.cells?.[p.colDef.colId ?? ''];
    const st = cellData?.style;
    const rowId = p.data?.id;
    const r = typeof rowId === 'number' ? rowId - 1 : (p.rowIndex ?? 0);
    const c = /^\d+$/.test(p.colDef.colId ?? '') ? Number(p.colDef.colId) : -1;

    // Right-align numbers by default if no explicit alignment
    if (cellData && typeof cellData.v === 'number' && !st?.hAlign) {
      out.textAlign = 'right';
      out.fontFamily = '"JetBrains Mono", monospace';
      out.fontFeatureSettings = '"tnum" 1';
    }

    // Style errors (#REF!, #DIV/0!, #CYCLE!, etc.)
    if (cellData && typeof cellData.text === 'string' && cellData.text.startsWith('#')) {
      out.color = currentPalette.error;
      out.fontWeight = '600';
      out.fontFamily = '"JetBrains Mono", monospace';
    }

    if (st) {
      if (st.bold) out.fontWeight = '700';
      if (st.italic) out.fontStyle = 'italic';
      if (st.underline) out.textDecoration = 'underline';
      if (st.fontName) out.fontFamily = st.fontName;
      if (st.fontSize) out.fontSize = `${st.fontSize}px`;
      if (st.color) out.color = st.color;
      if (st.bg) out.backgroundColor = st.bg;
      if (st.hAlign) out.textAlign = st.hAlign;
      if (st.vAlign) out.verticalAlign = st.vAlign === 'middle' ? 'middle' : st.vAlign;
      if (st.wrap) { out.whiteSpace = 'normal'; out.lineHeight = '1.25'; }
      if (st.border) {
        const b = `1px solid ${currentPalette.borderStrong}`;
        if (st.border.top) out.borderTop = b;
        if (st.border.bottom) out.borderBottom = b;
        if (st.border.left) out.borderLeft = b;
        if (st.border.right) out.borderRight = b;
      }
    }

    // Conditional Formatting evaluation
    if (cfRules.length > 0 && sheetName && c >= 0) {
      const cf = evaluateConditionalFormatting(
        cfRules,
        sheetName,
        r,
        c,
        cellData?.v ?? null,
        (rg) => getRangeStats(sheetName, rg)
      );
      if (cf) {
        if (cf.bg) out.backgroundColor = cf.bg;
        if (cf.color) out.color = cf.color;
        if (cf.background) out.background = cf.background;
      }
    }

    // Selection styling
    const sel = selRef.current;
    if (sel && c >= 0) {
      const inSel = sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);

      if (inSel && !(sel.active.r === r && sel.active.c === c)) {
        out.backgroundColor = currentPalette.accentSubtle;
      }

      if (sel.active.r === r && sel.active.c === c) {
        out.outline = `2px solid ${currentPalette.primary}`;
        out.outlineOffset = '-1px';
        out.zIndex = '3';
      }
    }

    // Formula references color bounding boxes
    if (formulaHighlights.length > 0 && c >= 0) {
      for (const h of formulaHighlights) {
        const rg = h.range;
        if (r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) {
          out.backgroundColor = `${h.color}15`;
          if (r === rg.r1) out.borderTop = `2px solid ${h.color}`;
          if (r === rg.r2) out.borderBottom = `2px solid ${h.color}`;
          if (c === rg.c1) out.borderLeft = `2px solid ${h.color}`;
          if (c === rg.c2) out.borderRight = `2px solid ${h.color}`;
        }
      }
    }

    // Marching ants for clipboard range (Ctrl+C / Ctrl+X)
    if (clipboardRange && sheetName && clipboardRange.sheet.toLowerCase() === sheetName.toLowerCase() && c >= 0) {
      const cr = clipboardRange.range;
      if (r >= cr.r1 && r <= cr.r2 && c >= cr.c1 && c <= cr.c2) {
        const dashed = `2px dashed ${currentPalette.primary}`;
        if (r === cr.r1) out.borderTop = dashed;
        if (r === cr.r2) out.borderBottom = dashed;
        if (c === cr.c1) out.borderLeft = dashed;
        if (c === cr.c2) out.borderRight = dashed;
      }
    }

    return out;
  }, [currentPalette, selRef, cfRules, sheetName, formulaHighlights, clipboardRange, getRangeStats]);

  // Column definitions
  const columns = useMemo<ColDef<Row>[]>(() => {
    const defs: ColDef<Row>[] = [{
      headerName: '#',
      colId: 'row',
      valueGetter: (p) => p.data?.id,
      pinned: 'left',
      width: 58,
      sortable: false,
      filter: false,
      resizable: false,
      editable: false,
      suppressMovable: true,
      cellRenderer: RowHeader,
      cellStyle: {
        background: currentPalette.gridRowHeaderBg,
        fontWeight: 600,
        textAlign: 'center',
        padding: '0',
        color: currentPalette.textSecondary,
        fontFamily: 'monospace',
      },
    }];

    const widths = model && sheetName ? model.getSheet(sheetName)?.colWidths : undefined;
    const hidden = model && sheetName ? model.getSheet(sheetName)?.hiddenCols : undefined;

    for (let c = 0; c < colCount; c++) {
      if (hidden?.has(c)) continue;
      defs.push({
        headerName: colToLetter(c),
        colId: String(c),
        valueGetter: (p) => {
          const cell = p.data?.cells?.[String(c)];
          if (!cell) return null;
          return cell.formula !== undefined ? '=' + cell.formula : cell.v ?? null;
        },
        valueFormatter: (p) => p.data?.cells?.[String(c)]?.text ?? '',
        valueSetter: (p) => {
          try {
            const raw = p.newValue;
            const text = raw === null || raw === undefined ? '' : String(raw);
            const st = useViewer.getState();
            if (!st.model || !st.sheetName || st.busy) return false;
            if (!p.data) return false;
            const r = p.data.id - 1;
            const cur = st.model.displayOf(st.sheetName, r, c);
            const curEdit = cur.formula !== undefined ? '=' + cur.formula : cur.text;
            if (text === curEdit) return false;
            const change = inputToChange(text, st.model.getRec(st.sheetName, r, c)?.style?.numFmt === '@');
            st.exec(editCellsCommand(st.model, st.sheetName, [{ r, c, next: change }], 'Edit cell'));
            const d = st.model.displayOf(st.sheetName, r, c);
            p.data.cells[String(c)] = {
              v: d.v,
              text: d.text,
              ...(d.formula !== undefined ? { formula: d.formula } : {}),
              ...(st.model.getRec(st.sheetName, r, c)?.style ? { style: st.model.getRec(st.sheetName, r, c)!.style } : {})
            };
            st.setSelection({ active: { r, c }, ranges: [{ r1: r, c1: c, r2: r, c2: c }] });
            anchorRef.current = { r, c };
            return true;
          } catch (e) {
            useViewer.getState().set({ error: e instanceof Error ? e.message : String(e) });
            return false;
          }
        },
        filter: 'agTextColumnFilter',
        filterValueGetter: (p) => p.data?.cells?.[String(c)]?.text ?? '',
        getQuickFilterText: (p) => p.data?.cells?.[String(c)]?.text ?? '',
        width: Math.round((widths?.has(c) ? widths.get(c)! * 7 + 5 : 140) * s.zoom / 100),
        minWidth: 40,
        sortable: true,
        resizable: true,
        editable: () => !useViewer.getState().busy,
        suppressMovable: true,
        cellStyle: cellStyleFn,
        comparator: (a, b, nodeA, nodeB) => {
          a = nodeA.data?.cells[String(c)]?.v ?? null;
          b = nodeB.data?.cells[String(c)]?.v ?? null;
          if (a == null) return b == null ? 0 : -1;
          if (b == null) return 1;
          const na = typeof a === 'number' ? a : (typeof a === 'string' && a !== '' && isFinite(Number(a)) ? Number(a) : null);
          const nb = typeof b === 'number' ? b : (typeof b === 'string' && b !== '' && isFinite(Number(b)) ? Number(b) : null);
          if (na !== null && nb !== null) return na - nb;
          return String(a).localeCompare(String(b), undefined, { numeric: true });
        },
      });
    }

    return defs.map((d) => (/^\d+$/.test(d.colId!) && Number(d.colId!) < frozenCols ? { ...d, pinned: 'left' as const } : d));
  }, [colCount, model, sheetName, frozenCols, cellStyleFn, s.layoutRev, s.zoom, currentPalette, anchorRef]);

  const gridTheme = useMemo(() => getGridTheme(mode, s.zoom), [mode, s.zoom]);

  const pinnedTop = useMemo(() => {
    if (!model || !sheetName) return undefined;
    const n = model.getSheet(sheetName)?.frozenRows ?? 0;
    if (!n) return undefined;
    return s.rows.filter(row => row.id <= n);
  }, [model, sheetName, s.rows]);

  const getRowHeight = useCallback((p: RowHeightParams<Row>) => {
    const st = useViewer.getState();
    const h = st.model && st.sheetName ? st.model.getSheet(st.sheetName!)?.rowHeights.get((p.data?.id ?? 1) - 1) : undefined;
    return (h ?? 30) * st.zoom / 100;
  }, []);

  useEffect(() => {
    const api = gridRef.current;
    if (!api || api.isDestroyed()) return;
    api.resetRowHeights();
    for (let i = 0; i < api.getPinnedTopRowCount(); i++) {
      const node = api.getPinnedTopRow(i);
      if (node?.data && model && sheetName) {
        node.setRowHeight((model.getSheet(sheetName)?.rowHeights.get(node.data.id - 1) ?? 30) * s.zoom / 100);
      }
    }
    api.onRowHeightChanged();
  }, [s.layoutRev, s.zoom, model, sheetName, gridRef]);

  // Update fill handle position whenever selection or scroll changes
  const updateFillHandle = useCallback(() => {
    const sel = selRef.current;
    if (!sel || isFormulaEditing) {
      setHandlePos(null);
      return;
    }
    const rg = sel.ranges[0] ?? { r1: sel.active.r, c1: sel.active.c, r2: sel.active.r, c2: sel.active.c };
    const cellEl = document.querySelector(`.ag-row[row-index="${rg.r2}"] [col-id="${rg.c2}"]`);
    const containerEl = containerRef.current;
    if (cellEl && containerEl) {
      const cRect = cellEl.getBoundingClientRect();
      const gRect = containerEl.getBoundingClientRect();
      setHandlePos({
        top: cRect.bottom - gRect.top - 4,
        left: cRect.right - gRect.left - 4,
        r1: rg.r1,
        c1: rg.c1,
        r2: rg.r2,
        c2: rg.c2,
      });
    } else {
      setHandlePos(null);
    }
  }, [selRef, isFormulaEditing]);

  useEffect(() => {
    updateFillHandle();
  }, [s.selection, s.zoom, s.rev, updateFillHandle]);

  // Handle Drag-to-Fill
  const startFillDrag = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!handlePos || !model || !sheetName) return;

    let targetR = handlePos.r2;
    let targetC = handlePos.c2;

    const onMouseMove = (ev: MouseEvent) => {
      const el = document.elementFromPoint(ev.clientX, ev.clientY);
      const cellEl = el?.closest('.ag-cell') as HTMLElement | null;
      const rowEl = el?.closest('.ag-row') as HTMLElement | null;
      if (!cellEl || !rowEl) return;

      const rAttr = rowEl.getAttribute('row-index');
      const cAttr = cellEl.getAttribute('col-id');
      if (!rAttr || !cAttr || !/^\d+$/.test(cAttr)) return;

      const r = parseInt(rAttr, 10);
      const c = parseInt(cAttr, 10);

      // Support downward or rightward fill
      if (r > handlePos.r2) {
        targetR = r;
        targetC = handlePos.c2;
      } else if (c > handlePos.c2) {
        targetC = c;
        targetR = handlePos.r2;
      }

      // Compute preview box
      const startCell = document.querySelector(`.ag-row[row-index="${handlePos.r1}"] [col-id="${handlePos.c1}"]`);
      const endCell = document.querySelector(`.ag-row[row-index="${targetR}"] [col-id="${targetC}"]`);
      const contEl = containerRef.current;
      if (startCell && endCell && contEl) {
        const sRect = startCell.getBoundingClientRect();
        const eRect = endCell.getBoundingClientRect();
        const cRect = contEl.getBoundingClientRect();
        setDragPreview({
          top: sRect.top - cRect.top,
          left: sRect.left - cRect.left,
          width: eRect.right - sRect.left,
          height: eRect.bottom - sRect.top,
        });
      }
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      setDragPreview(null);

      if (!model || !sheetName) return;
      const st = useViewer.getState();
      const changes: { r: number; c: number; next: ReturnType<typeof inputToChange> }[] = [];

      // Downward fill
      if (targetR > handlePos.r2) {
        for (let col = handlePos.c1; col <= handlePos.c2; col++) {
          const srcCells: { r: number; val: unknown; formula?: string; numFmt?: string }[] = [];
          for (let row = handlePos.r1; row <= handlePos.r2; row++) {
            const disp = model.displayOf(sheetName, row, col);
            const rec = model.getRec(sheetName, row, col);
            srcCells.push({ r: row, val: disp.v, formula: disp.formula, numFmt: rec?.style?.numFmt });
          }

          const hasFormula = srcCells.some(sc => sc.formula !== undefined);
          const allNums = srcCells.every(sc => typeof sc.val === 'number');

          for (let row = handlePos.r2 + 1; row <= targetR; row++) {
            if (hasFormula) {
              const patternIdx = (row - (handlePos.r2 + 1)) % srcCells.length;
              const src = srcCells[patternIdx];
              if (src.formula) {
                const deltaRow = row - src.r;
                const shifted = translateFormula('=' + src.formula, deltaRow, 0);
                changes.push({ r: row, c: col, next: inputToChange(shifted) });
              } else {
                changes.push({ r: row, c: col, next: { v: src.val as string | number | boolean | null } });
              }
            } else if (allNums && srcCells.length >= 2) {
              const v0 = srcCells[0].val as number;
              const v1 = srcCells[srcCells.length - 1].val as number;
              const step = (v1 - v0) / (srcCells.length - 1);
              const nextVal = v1 + step * (row - handlePos.r2);
              changes.push({ r: row, c: col, next: { v: nextVal } });
            } else if (allNums && srcCells.length === 1) {
              changes.push({ r: row, c: col, next: { v: srcCells[0].val as number } });
            } else {
              const patternIdx = (row - (handlePos.r2 + 1)) % srcCells.length;
              changes.push({ r: row, c: col, next: { v: srcCells[patternIdx].val as string | number | boolean | null } });
            }
          }
        }
      } else if (targetC > handlePos.c2) {
        // Rightward fill
        for (let row = handlePos.r1; row <= handlePos.r2; row++) {
          const srcCells: { c: number; val: unknown; formula?: string; numFmt?: string }[] = [];
          for (let col = handlePos.c1; col <= handlePos.c2; col++) {
            const disp = model.displayOf(sheetName, row, col);
            const rec = model.getRec(sheetName, row, col);
            srcCells.push({ c: col, val: disp.v, formula: disp.formula, numFmt: rec?.style?.numFmt });
          }

          const hasFormula = srcCells.some(sc => sc.formula !== undefined);
          const allNums = srcCells.every(sc => typeof sc.val === 'number');

          for (let col = handlePos.c2 + 1; col <= targetC; col++) {
            if (hasFormula) {
              const patternIdx = (col - (handlePos.c2 + 1)) % srcCells.length;
              const src = srcCells[patternIdx];
              if (src.formula) {
                const deltaCol = col - src.c;
                const shifted = translateFormula('=' + src.formula, 0, deltaCol);
                changes.push({ r: row, c: col, next: inputToChange(shifted) });
              } else {
                changes.push({ r: row, c: col, next: { v: src.val as string | number | boolean | null } });
              }
            } else if (allNums && srcCells.length >= 2) {
              const v0 = srcCells[0].val as number;
              const v1 = srcCells[srcCells.length - 1].val as number;
              const step = (v1 - v0) / (srcCells.length - 1);
              const nextVal = v1 + step * (col - handlePos.c2);
              changes.push({ r: row, c: col, next: { v: nextVal } });
            } else if (allNums && srcCells.length === 1) {
              changes.push({ r: row, c: col, next: { v: srcCells[0].val as number } });
            } else {
              const patternIdx = (col - (handlePos.c2 + 1)) % srcCells.length;
              changes.push({ r: row, c: col, next: { v: srcCells[patternIdx].val as string | number | boolean | null } });
            }
          }
        }
      }

      if (changes.length > 0) {
        st.exec(editCellsCommand(st.model!, st.sheetName!, changes, 'Auto fill'));
        st.setSelection({
          active: { r: handlePos.r1, c: handlePos.c1 },
          ranges: [{ r1: handlePos.r1, c1: handlePos.c1, r2: targetR, c2: targetC }],
        });
      }
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const onColumnResized = useCallback((event: ColumnResizedEvent<Row>) => {
    if (!event.finished || !['uiColumnResized', 'autosizeColumns'].includes(event.source)) return;
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName) return;
    const sh = st.model.getSheet(st.sheetName)!;
    const sizes = new Map<number, number>();
    for (const column of event.columns ?? (event.column ? [event.column] : [])) {
      if (!/^\d+$/.test(column.getColId())) continue;
      const c = Number(column.getColId());
      const width = Math.max(1, (column.getActualWidth() / (st.zoom / 100) - 5) / 7);
      if (Math.abs((sh.colWidths.get(c) ?? (140 - 5) / 7) - width) > 0.05) sizes.set(c, width);
    }
    if (!sizes.size) return;
    queueMicrotask(() => {
      const current = useViewer.getState();
      if (current.busy || current.model !== st.model || current.sheetName !== st.sheetName) return;
      current.exec(resizeCommand(st.model!, st.sheetName!, 'col', sizes));
    });
  }, []);

  const openCellMenu = (event: CellContextMenuEvent<Row>) => {
    if (!(event.event instanceof MouseEvent) || !event.data || s.busy) return;
    event.event.preventDefault();
    const r = event.data.id - 1;
    const column = event.column.getColId();
    const c = /^\d+$/.test(column) ? Number(column) : 0;
    const st = useViewer.getState();
    const selected = event.api.getSelectedNodes().flatMap(n => n.data ? [n.data.id - 1] : []);
    const ranges = st.selection?.ranges ?? [];
    const inRange = ranges.some(range => r >= range.r1 && r <= range.r2 && (column === 'row' || (c >= range.c1 && c <= range.c2)));
    const rows = selected.includes(r) ? selected.map(start => ({ start, count: 1 })) : inRange ? ranges.map(range => ({ start: range.r1, count: range.r2 - range.r1 + 1 })) : [{ start: r, count: 1 }];
    const cols = inRange ? ranges.map(range => ({ start: range.c1, count: range.c2 - range.c1 + 1 })) : [{ start: c, count: 1 }];
    if (!inRange) st.setSelection({ active: { r, c }, ranges: [{ r1: r, r2: r, c1: c, c2: c }] });
    setContext({ x: event.event.clientX, y: event.event.clientY, rows: mergeSpans(rows), cols: mergeSpans(cols) });
  };

  const openHeaderMenu = (event: React.MouseEvent) => {
    const header = (event.target as HTMLElement).closest('.ag-header-cell');
    const id = header?.getAttribute('col-id');
    if (!id || !/^\d+$/.test(id) || s.busy) return;
    event.preventDefault();
    const c = Number(id);
    useViewer.getState().setSelection({ active: { r: 0, c }, ranges: [{ r1: 0, r2: 0, c1: c, c2: c }] });
    setContext({ x: event.clientX, y: event.clientY, rows: [], cols: [{ start: c, count: 1 }] });
  };

  const contextOperation = (axis: 'row' | 'col', opMode: 'insert' | 'delete') => {
    const st = useViewer.getState();
    const spans = axis === 'row' ? context?.rows : context?.cols;
    setContext(null);
    if (st.busy || !st.model || !st.sheetName || !spans?.length) return;
    st.exec(axisCommand(st.model, st.sheetName, axis, opMode, spans));
    st.setSelection(null);
  };

  const contextSize = (axis: 'row' | 'col') => {
    const st = useViewer.getState();
    if (!context || !st.model || !st.sheetName) return;
    const spans = axis === 'row' ? context.rows : context.cols;
    if (!spans.length) return;
    const start = spans[0].start;
    const sh = st.model.getSheet(st.sheetName)!;
    st.setSelection({
      active: axis === 'row' ? { r: start, c: 0 } : { r: 0, c: start },
      ranges: spans.map(span => axis === 'row'
        ? { r1: span.start, r2: span.start + span.count - 1, c1: 0, c2: 0 }
        : { r1: 0, r2: 0, c1: span.start, c2: span.start + span.count - 1 }
      )
    });
    const val = String(axis === 'row' ? sh.rowHeights.get(start) ?? 30 : sh.colWidths.get(start) ?? (140 - 5) / 7);
    onOpenDimWithValues(axis === 'row' ? 'rowHeight' : 'colWidth', val);
    setContext(null);
  };

  const onCellFocused = (e: CellFocusedEvent) => {
    if (e.rowIndex == null || !e.column) return;
    const col = typeof e.column === 'string' ? e.column : e.column.getColId();
    if (!/^\d+$/.test(col)) return;
    const row = (e.rowPinned === 'top' ? e.api.getPinnedTopRow(e.rowIndex) : e.api.getDisplayedRowAtIndex(e.rowIndex))?.data as Row | undefined;
    if (!row) return;
    const r = (row.id as number) - 1;
    const cc = Number(col);
    const st = useViewer.getState();

    if (isFormulaEditing) return;

    if (shiftDownRef.current && anchorRef.current && st.selection) {
      const base = anchorRef.current;
      st.setSelection({ active: { r, c: cc }, ranges: [normRange(base, { r, c: cc })] });
    } else if (!st.selection || st.selection.active.r !== r || st.selection.active.c !== cc) {
      anchorRef.current = { r, c: cc };
      st.setSelection({ active: { r, c: cc }, ranges: [{ r1: r, c1: cc, r2: r, c2: cc }] });
    }
    gridRef.current?.refreshCells({ force: true, columns: undefined });
  };

  // Cell click: Handles formula reference insertion, single cell click, and ctrl+click
  const onCellClicked = (e: CellClickedEvent) => {
    const colId = typeof e.column === 'string' ? e.column : e.column.getColId();
    const row = e.data as Row | undefined;
    if (!row) return;
    const r = (row.id as number) - 1;
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;

    if (colId === 'row') {
      const sh = st.model.getSheet(st.sheetName)!;
      onSelect({ r, c: 0 }, [{ r1: r, c1: 0, r2: r, c2: Math.max(0, sh.colCount - 1) }]);
      return;
    }

    if (!/^\d+$/.test(colId)) return;
    const c = Number(colId);

    // Interactive formula reference click insertion
    if (isFormulaEditing && onFormulaRangeSelect) {
      if (shiftDownRef.current && anchorRef.current) {
        const base = anchorRef.current;
        const norm = normRange(base, { r, c });
        const rangeStr = `${toAddress(norm.r1, norm.c1)}:${toAddress(norm.r2, norm.c2)}`;
        onFormulaRangeSelect(rangeStr);
      } else {
        anchorRef.current = { r, c };
        onFormulaRangeSelect(toAddress(r, c));
      }
      return;
    }

    if (e.event instanceof MouseEvent && (e.event.ctrlKey || e.event.metaKey) && st.selection) {
      onSelect({ r, c }, [...st.selection.ranges, { r1: r, c1: c, r2: r, c2: c }]);
    } else {
      anchorRef.current = { r, c };
      st.setSelection({ active: { r, c }, ranges: [{ r1: r, c1: c, r2: r, c2: c }] });
    }
  };

  const onHeaderClick = (e: React.MouseEvent) => {
    if ((e.target as HTMLElement).closest('.ag-header-cell-resize,.ag-header-cell-filter-button,.ag-header-cell-menu-button,.ag-header-cell-text,.ag-header-cell-label')) return;
    const el = (e.target as HTMLElement).closest('.ag-header-cell');
    if (!el) return;
    const colId = el.getAttribute('col-id');
    if (!colId || !/^\d+$/.test(colId)) return;
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const c = Number(colId);
    const sh = st.model.getSheet(st.sheetName)!;
    onSelect({ r: 0, c }, [{ r1: 0, c1: c, r2: Math.max(0, sh.rowCount - 1), c2: c }]);
  };

  return (
    <main
      ref={containerRef}
      className="grid"
      aria-busy={s.busy}
      onClick={onHeaderClick}
      onContextMenu={openHeaderMenu}
      style={{ position: 'relative' }}
    >
      {sheetName && (
        <AgGridReact<Row>
          key={(s.book?.name ?? '') + (sheetName ?? '')}
          theme={gridTheme}
          rowData={pinnedTop?.length ? s.rows.filter(row => !pinnedTop.some(p => p.id === row.id)) : s.rows}
          columnDefs={columns}
          defaultColDef={{ editable: false }}
          getRowId={(p) => String(p.data.id)}
          quickFilterText={s.search}
          pinnedTopRowData={pinnedTop as Row[] | undefined}
          getRowHeight={getRowHeight}
          stopEditingWhenCellsLoseFocus
          rowSelection={{ mode: "multiRow", enableClickSelection: false }}
          onColumnResized={onColumnResized}
          onCellContextMenu={openCellMenu}
          preventDefaultOnContextMenu
          onGridReady={(e) => { gridRef.current = e.api; }}
          onCellFocused={onCellFocused}
          onCellClicked={onCellClicked}
          onBodyScroll={updateFillHandle}
          overlayNoRowsTemplate="No matching rows"
        />
      )}

      {/* Excel Drag Fill Handle */}
      {handlePos && !isFormulaEditing && (
        <div
          className="cs-fill-handle"
          title="Drag to fill formula or values"
          style={{
            top: handlePos.top,
            left: handlePos.left,
          }}
          onMouseDown={startFillDrag}
        />
      )}

      {/* Drag Fill Preview Outline */}
      {dragPreview && (
        <div
          className="cs-fill-preview"
          style={{
            top: dragPreview.top,
            left: dragPreview.left,
            width: dragPreview.width,
            height: dragPreview.height,
          }}
        />
      )}

      {s.busy && (
        <div className="loading">
          <div className="loading-spinner" />
          <span>Reading spreadsheet…</span>
        </div>
      )}

      {/* Header and Cell context menu */}
      <Menu
        open={!!context}
        onClose={() => setContext(null)}
        anchorReference="anchorPosition"
        anchorPosition={context ? { top: context.y, left: context.x } : undefined}
      >
        {!!context?.rows.length && <MenuItem onClick={() => contextOperation('row', 'insert')}>Insert row(s) above</MenuItem>}
        {!!context?.rows.length && <MenuItem onClick={() => contextOperation('row', 'delete')}>Delete row(s)</MenuItem>}
        {!!context?.rows.length && <MenuItem onClick={() => contextSize('row')}>Row height…</MenuItem>}
        {!!context?.rows.length && <Divider />}
        <MenuItem onClick={() => contextOperation('col', 'insert')}>Insert column(s) before</MenuItem>
        <MenuItem onClick={() => contextOperation('col', 'delete')}>Delete column(s)</MenuItem>
        <MenuItem onClick={() => contextSize('col')}>Column width…</MenuItem>
      </Menu>
    </main>
  );
};
