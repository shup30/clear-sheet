import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  ThemeProvider, createTheme, CssBaseline, Button, TextField, Typography, Box, Chip,
  Alert, LinearProgress, Tabs, Tab, MenuItem, Select, Snackbar, Toolbar as MuiToolbar,
  IconButton, Tooltip, Divider, Menu, Dialog, DialogTitle, DialogContent, DialogActions,
  Checkbox, FormControlLabel, List, ListItemButton, InputLabel, FormControl,
} from '@mui/material';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, ModuleRegistry, themeQuartz, type ColDef, type GridApi, type CellFocusedEvent, type CellClickedEvent, type CellClassParams, type RowHeightParams, type ColumnResizedEvent, type CellContextMenuEvent } from 'ag-grid-community';
import type { Row, CellStyle } from '../electron/types';
import { useViewer, type SelRange } from './store';
import { colToLetter } from './workbook/cellRef';
import { FONT_CHOICES, SIZE_CHOICES, NUM_FORMATS } from './workbook/format';
import { editCellsCommand, formatCommand, insertDeleteCommand, sheetOpCommand, inputToChange, layoutCommand, compositeCommand } from './workbook/commands';
import { translateFormula } from './workbook/formulas';
import { rangeToTSV, parseTSV } from './workbook/clipboard';
import { RowHeader } from './RowHeader';
import { resizeCommand } from './workbook/dimensions';
import { axisCommand, mergeSpans, type Span } from './workbook/axisOperations';
import { DESKTOP_UPDATE_MESSAGE } from './desktop';
import './style.css';

ModuleRegistry.registerModules([AllCommunityModule]);

const theme = createTheme({
  palette: { primary: { main: '#16744b' }, background: { default: '#f7f9fb' } },
  typography: { fontFamily: 'Segoe UI, sans-serif', button: { textTransform: 'none' } },
  shape: { borderRadius: 8 },
});

function normRange(a: { r: number; c: number }, b: { r: number; c: number }): SelRange {
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}

function rangeLabel(r: SelRange): string {
  const a = `${colToLetter(r.c1)}${r.r1 + 1}`;
  const b = `${colToLetter(r.c2)}${r.r2 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

function targetsOf(ranges: SelRange[]): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  if(ranges.reduce((sum,rg)=>sum+(rg.r2-rg.r1+1)*(rg.c2-rg.c1+1),0)>200000){useViewer.getState().set({error:'This operation is limited to 200,000 cells.'});return out;}
  for (const rg of ranges) {
    const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    if (area > 200000) continue;
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push({ r, c });
  }
  return out;
}

interface Found { sheet: string; r: number; c: number; text: string }

function App() {
  const s = useViewer();
  const grid = useRef<GridApi<Row> | null>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const anchor = useRef<{ r: number; c: number } | null>(null);
  const shiftDown = useRef(false);
  const localClipboard=useRef<{text:string;cut:boolean;r:number;c:number;cells:{r:number;c:number;next:{v:string|number|boolean|null;f?:string}}[]}|null>(null);
  const selRef = useRef(s.selection);
  selRef.current = s.selection;
  const [toast, setToast] = useState('');
  const [drag, setDrag] = useState(false);
  const [query, setQuery] = useState('');
  const [fb, setFb] = useState('');
  const [fbEditing, setFbEditing] = useState(false);
  const [freezeMenu, setFreezeMenu] = useState<null | HTMLElement>(null);
  const [insertMenu, setInsertMenu] = useState<null | HTMLElement>(null);
  const [sheetMenu, setSheetMenu] = useState<null | HTMLElement>(null);
  const [findOpen, setFindOpen] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [findScope, setFindScope] = useState<'sheet' | 'book'>('sheet');
  const [matchCase, setMatchCase] = useState(false);
  const [found, setFound] = useState<Found[]>([]);
  const [foundIdx, setFoundIdx] = useState(-1);
  const [renameSheet, setRenameSheet] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState('');
  const [deleteSheet, setDeleteSheet] = useState<string | null>(null);
  const [dimKind, setDimKind] = useState<'colWidth' | 'rowHeight' | null>(null);
  const [dimValue, setDimValue] = useState('');
  const [context, setContext] = useState<{x:number;y:number;rows:Span[];cols:Span[]}|null>(null);

  const sheetName = s.sheetName;
  const model = s.model;

  useEffect(() => { void s.refresh(); return window.viewer.onOpen(() => void useViewer.getState().open()); }, []);
  useEffect(() => { const t = setTimeout(() => s.set({ search: query }), 180); return () => clearTimeout(t); }, [query]);
  useEffect(() => { setQuery('');setContext(null); }, [s.sheetName, s.book?.name]);
  useEffect(() => {
    if (!fbEditing) setFb(s.selected?.formula !== undefined ? '=' + s.selected.formula : s.selected?.text ?? '');
  }, [s.selected, fbEditing]);
  useEffect(() => {
    const before = (e: BeforeUnloadEvent) => { if (useViewer.getState().dirty || useViewer.getState().busy) { e.preventDefault(); e.returnValue = ''; } };
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
    const api=grid.current;
    if(!api)return;
    const pinned=Array.from({length:api.getPinnedTopRowCount()},(_,i)=>api.getPinnedTopRow(i)).find(n=>n?.data?.id===r+1);
    const node=pinned ?? api.getRowNode(String(r+1));
    if(node?.rowIndex==null)return;
    if(!pinned)api.ensureIndexVisible(node.rowIndex);
    api.ensureColumnVisible(String(c));
    api.setFocusedCell(node.rowIndex,String(c),pinned?'top':null);
  }, []);

  const select = useCallback((active: { r: number; c: number }, ranges: SelRange[]) => {
    anchor.current = active;
    useViewer.getState().setSelection({ active, ranges });
    focusCell(active.r, active.c);
  }, [focusCell]);

  const gotoFound = useCallback((f: Found) => {
    const st = useViewer.getState();
    if (f.sheet !== st.sheetName) { void st.load(f.sheet).then(() => select({ r: f.r, c: f.c }, [normRange({ r: f.r, c: f.c }, { r: f.r, c: f.c })])); }
    else select({ r: f.r, c: f.c }, [normRange({ r: f.r, c: f.c }, { r: f.r, c: f.c })]);
  }, [select]);

  // ---------- clipboard ----------

  const doCopy = useCallback(async (cut: boolean) => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) { setToast('Select a cell first'); return; }
    const ranges = st.selection.ranges.length ? st.selection.ranges : [normRange(st.selection.active, st.selection.active)];
    if(ranges.reduce((n,rg)=>n+(rg.r2-rg.r1+1)*(rg.c2-rg.c1+1),0)>200000){st.set({error:'Copy/cut is limited to 200,000 cells.'});return;}
    const parts = ranges.map((rg) => rangeToTSV(st.model!, st.sheetName!, rg.r1, rg.c1, rg.r2, rg.c2, true));
    const clipboardText=parts.join('\r\n');
    const rg=ranges[0];
    const cells=ranges.length===1?targetsOf(ranges).map(({r,c})=>{const rec=st.model!.getRec(st.sheetName!,r,c);return {r,c,next:{v:rec?.v??null,f:rec?.f}};}):[];
    const copied=await copyText(clipboardText, cut ? 'Cut to clipboard' : 'Copied to clipboard');
    if(!copied || useViewer.getState().model!==st.model || useViewer.getState().rev!==st.rev || useViewer.getState().busy)return;
    localClipboard.current=cells.length?{text:clipboardText,cut,r:rg.r1,c:rg.c1,cells}:null;
    if (cut) {
      const targets = targetsOf(ranges);
      if (targets.length) {
        st.exec(editCellsCommand(st.model, st.sheetName, targets.map((t) => ({ r: t.r, c: t.c, next: { v: null } })), 'Cut'));
      }
    }
  }, []);

  const doPaste = useCallback(async () => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) { setToast('Select a cell first'); return; }
    let text = '';
    try {
      const r = await window.viewer.paste();
      if (!r.ok) throw new Error(r.error);
      text = r.data;
    } catch {
      try { text = await navigator.clipboard.readText(); } catch { setToast('Clipboard is empty or unavailable'); return; }
    }
    if(useViewer.getState().model!==st.model || useViewer.getState().rev!==st.rev || useViewer.getState().busy)return;
    if (!text) { setToast('Clipboard is empty'); return; }
    const gridData = parseTSV(text);
    if (!gridData.length) return;
    const { r: ar, c: ac } = st.selection.active;
    let changes = gridData.flatMap((row, dr) => row.map((t, dc) => ({ r: ar + dr, c: ac + dc, next: inputToChange(t) })));
    const local=localClipboard.current;
    if(local?.text===text)changes=local.cells.map(cell=>({r:ar+cell.r-local.r,c:ac+cell.c-local.c,next:{...cell.next,f:cell.next.f===undefined?undefined:local.cut?cell.next.f:translateFormula(cell.next.f,ar-local.r,ac-local.c)}}));
    if (changes.length > 500000) { st.set({ error: 'Paste is limited to 500,000 cells.' }); return; }
    if(changes.some(c=>c.r>=1048576 || c.c>=16384)){st.set({error:'Paste exceeds Excel row/column limits.'});return;}
    st.exec(editCellsCommand(st.model, st.sheetName, changes, 'Paste'));
    const last = changes[changes.length - 1];
    select({ r: last.r, c: last.c }, [{ r1: ar, c1: ac, r2: last.r, c2: last.c }]);
  }, [select]);

  const doDelete = useCallback(() => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const ranges = st.selection.ranges.length ? st.selection.ranges : [normRange(st.selection.active, st.selection.active)];
    const targets = targetsOf(ranges);
    if (!targets.length) return;
    st.exec(editCellsCommand(st.model, st.sheetName, targets.map((t) => ({ r: t.r, c: t.c, next: { v: null } })), 'Clear cells'));
  }, []);

  // ---------- global shortcuts ----------

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const st = useViewer.getState();
      shiftDown.current = e.shiftKey;
      const target = e.target as HTMLElement;
      const inField = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target.isContentEditable;
      const gridEditing = !!document.querySelector('.ag-cell-inline-editing');
      if (e.key === 'Shift') shiftDown.current = true;
      if(st.busy)return;
      if(target.closest('[role=dialog]'))return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f' && !e.shiftKey) { e.preventDefault(); searchRef.current?.focus(); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'h') { e.preventDefault(); setFindOpen(true); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        if (!st.model) return;
        grid.current?.stopEditing();
        if(document.activeElement instanceof HTMLInputElement) document.activeElement.blur();
        if (e.shiftKey) void useViewer.getState().saveAs(); else void useViewer.getState().save();
        return;
      }
      if (inField || gridEditing) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); st.undo(); return; }
      if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'y' || (e.key.toLowerCase() === 'z' && e.shiftKey))) { e.preventDefault(); st.redo(); return; }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        if (st.selection) { e.preventDefault(); void doCopy(false); } return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'x') {
        if (st.selection) { e.preventDefault(); void doCopy(true); } return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
        if (st.selection && st.model) { e.preventDefault(); void doPaste(); } return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
        if (st.model && st.sheetName) {
          e.preventDefault();
          const sh = st.model.getSheet(st.sheetName);
          if (sh) select({ r: 0, c: 0 }, [{ r1: 0, c1: 0, r2: Math.max(0, sh.rowCount - 1), c2: Math.max(0, sh.colCount - 1) }]);
        }
        return;
      }
      if ((e.key === 'Delete' || e.key === 'Backspace') && st.selection && st.model) { e.preventDefault(); doDelete(); }
    };
    const up = (e: KeyboardEvent) => { if (e.key === 'Shift') shiftDown.current = false; };
    window.addEventListener('keydown', handler);
    window.addEventListener('keyup', up);
    return () => { window.removeEventListener('keydown', handler); window.removeEventListener('keyup', up); };
  }, [doCopy, doPaste, doDelete, select]);

  // ---------- grid columns ----------

  const colCount = model && sheetName ? Math.min(Math.max(model.getSheet(sheetName)?.colCount ?? 10, 10), 1024) : 0;
  const frozenCols = model && sheetName ? model.getSheet(sheetName)?.frozenCols ?? 0 : 0;

  const cellStyleFn = useCallback((p: CellClassParams<Row>) => {
    const out: Record<string, string> = {fontWeight:'normal',fontStyle:'normal',textDecoration:'none',fontFamily:'inherit',fontSize:'inherit',color:'inherit',backgroundColor:'',textAlign:'left',verticalAlign:'middle',whiteSpace:'nowrap',lineHeight:'normal',borderTop:'',borderBottom:'',borderLeft:'',borderRight:'',outline:''};
    const st = p.data?.cells?.[p.colDef.colId ?? '']?.style;
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
        const b = '1px solid #555';
        if (st.border.top) out.borderTop = b;
        if (st.border.bottom) out.borderBottom = b;
        if (st.border.left) out.borderLeft = b;
        if (st.border.right) out.borderRight = b;
      }
    }
    const sel = selRef.current;
    if (sel && /^\d+$/.test(p.colDef.colId ?? '') && typeof p.rowIndex === 'number') {
      const rowId = p.data?.id;
      const r = (rowId ?? p.rowIndex + 1) - 1;
      const c = Number(p.colDef.colId);
      const inSel = sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2);
      if (inSel && !(sel.active.r === r && sel.active.c === c)) out.backgroundColor = out.backgroundColor ?? '#d7e9e0';
      if (sel.active.r === r && sel.active.c === c) out.outline = '2px solid #16744b';
    }
    return out;
  }, []);

  const columns = useMemo<ColDef<Row>[]>(() => {
    const defs: ColDef<Row>[] = [{
      headerName: '#', colId: 'row', valueGetter: (p) => p.data?.id, pinned: 'left', width: 64,
      sortable: false, filter: false, resizable: false, editable: false, suppressMovable: true,
      cellRenderer: RowHeader,
      cellStyle: { background: '#f0f3f1', fontWeight: 600, textAlign: 'center',padding:'0' },
    }];
    const widths = model && sheetName ? model.getSheet(sheetName)?.colWidths : undefined;
    const hidden = model && sheetName ? model.getSheet(sheetName)?.hiddenCols : undefined;
    for (let c = 0; c < colCount; c++) {
      if (hidden?.has(c)) continue;
      defs.push({
        headerName: colToLetter(c), colId: String(c),
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
            if(!p.data)return false;
            const r = p.data.id - 1;
            const cur = st.model.displayOf(st.sheetName, r, c);
            const curEdit = cur.formula !== undefined ? '=' + cur.formula : cur.text;
            if (text === curEdit) return false;
            const change = inputToChange(text,st.model.getRec(st.sheetName,r,c)?.style?.numFmt==='@');
            st.exec(editCellsCommand(st.model, st.sheetName, [{ r, c, next: change }], 'Edit cell'));
            const d = st.model.displayOf(st.sheetName, r, c);
            p.data.cells[String(c)] = { v: d.v, text: d.text, ...(d.formula !== undefined ? { formula: d.formula } : {}), ...(st.model.getRec(st.sheetName, r, c)?.style ? { style: st.model.getRec(st.sheetName, r, c)!.style } : {}) };
            st.setSelection({ active: { r, c }, ranges: [{ r1: r, c1: c, r2: r, c2: c }] });
            anchor.current = { r, c };
            return true;
          } catch(e) { useViewer.getState().set({error:e instanceof Error?e.message:String(e)});return false; }
        },
        filter: 'agTextColumnFilter',
        filterValueGetter: (p) => p.data?.cells?.[String(c)]?.text ?? '',
        getQuickFilterText: (p) => p.data?.cells?.[String(c)]?.text ?? '',
        width: Math.round((widths?.has(c) ? widths.get(c)! * 7 + 5 : 140) * s.zoom / 100), minWidth:40,
        sortable: true, resizable: true, editable: () => !useViewer.getState().busy, suppressMovable: true,
        cellStyle: cellStyleFn,
        comparator: (a, b, nodeA, nodeB) => {
          a=nodeA.data?.cells[String(c)]?.v??null;b=nodeB.data?.cells[String(c)]?.v??null;
          if (a == null) return b == null ? 0 : -1;
          if (b == null) return 1;
          const na = typeof a === 'number' ? a : (typeof a === 'string' && a !== '' && isFinite(Number(a)) ? Number(a) : null);
          const nb = typeof b === 'number' ? b : (typeof b === 'string' && b !== '' && isFinite(Number(b)) ? Number(b) : null);
          if (na !== null && nb !== null) return na - nb;
          return String(a).localeCompare(String(b), undefined, { numeric: true });
        },
      });
    }
    // frozen columns pin
    return defs.map((d) => (/^\d+$/.test(d.colId!) && Number(d.colId!) < frozenCols ? { ...d, pinned: 'left' as const } : d));
  }, [colCount, model, sheetName, frozenCols, cellStyleFn, s.layoutRev, s.zoom]);

  const gridTheme = useMemo(() => themeQuartz.withParams({
    accentColor: '#16744b', fontFamily: 'Segoe UI', fontSize: 13 * s.zoom / 100,
    rowHeight: 30 * s.zoom / 100, headerHeight: 34 * s.zoom / 100,
  }), [s.zoom]);

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
    const api=grid.current;
    if(!api || api.isDestroyed())return;
    api.resetRowHeights();
    for(let i=0;i<api.getPinnedTopRowCount();i++){
      const node=api.getPinnedTopRow(i);
      if(node?.data && model && sheetName)node.setRowHeight((model.getSheet(sheetName)?.rowHeights.get(node.data.id-1)??30)*s.zoom/100);
    }
    api.onRowHeightChanged();
  },[s.layoutRev,s.zoom,model,sheetName]);

  const onColumnResized = useCallback((event: ColumnResizedEvent<Row>) => {
    if(!event.finished || !['uiColumnResized','autosizeColumns'].includes(event.source))return;
    const st=useViewer.getState();
    if(st.busy || !st.model || !st.sheetName)return;
    const sh=st.model.getSheet(st.sheetName)!;
    const sizes=new Map<number,number>();
    for(const column of event.columns ?? (event.column ? [event.column] : [])){
      if(!/^\d+$/.test(column.getColId()))continue;
      const c=Number(column.getColId());
      const width=Math.max(1,(column.getActualWidth()/(st.zoom/100)-5)/7);
      if(Math.abs((sh.colWidths.get(c)??(140-5)/7)-width)>0.05)sizes.set(c,width);
    }
    if(!sizes.size)return;
    // AG Grid emits synchronously during a drag; update React after that finishes.
    queueMicrotask(()=>{
      const current=useViewer.getState();
      if(current.busy || current.model!==st.model || current.sheetName!==st.sheetName)return;
      current.exec(resizeCommand(st.model!,st.sheetName!,'col',sizes));
    });
  },[]);

  const openCellMenu = (event: CellContextMenuEvent<Row>) => {
    if(!(event.event instanceof MouseEvent) || !event.data || s.busy)return;
    event.event.preventDefault();
    const r=event.data.id-1;
    const column=event.column.getColId();
    const c=/^\d+$/.test(column)?Number(column):0;
    const st=useViewer.getState();
    const selected=event.api.getSelectedNodes().flatMap(n=>n.data?[n.data.id-1]:[]);
    const ranges=st.selection?.ranges ?? [];
    const inRange=ranges.some(range=>r>=range.r1&&r<=range.r2&&(column==='row'||(c>=range.c1&&c<=range.c2)));
    const rows=selected.includes(r)?selected.map(start=>({start,count:1})):inRange?ranges.map(range=>({start:range.r1,count:range.r2-range.r1+1})):[{start:r,count:1}];
    const cols=inRange?ranges.map(range=>({start:range.c1,count:range.c2-range.c1+1})):[{start:c,count:1}];
    if(!inRange)st.setSelection({active:{r,c},ranges:[{r1:r,r2:r,c1:c,c2:c}]});
    setContext({x:event.event.clientX,y:event.event.clientY,rows:mergeSpans(rows),cols:mergeSpans(cols)});
  };

  const openHeaderMenu = (event: React.MouseEvent) => {
    const header=(event.target as HTMLElement).closest('.ag-header-cell');
    const id=header?.getAttribute('col-id');
    if(!id || !/^\d+$/.test(id) || s.busy)return;
    event.preventDefault();
    const c=Number(id);
    useViewer.getState().setSelection({active:{r:0,c},ranges:[{r1:0,r2:0,c1:c,c2:c}]});
    setContext({x:event.clientX,y:event.clientY,rows:[],cols:[{start:c,count:1}]});
  };

  const contextOperation = (axis:'row'|'col',mode:'insert'|'delete') => {
    const st=useViewer.getState();const spans=axis==='row'?context?.rows:context?.cols;
    setContext(null);
    if(st.busy || !st.model || !st.sheetName || !spans?.length)return;
    st.exec(axisCommand(st.model,st.sheetName,axis,mode,spans));
    st.setSelection(null);
  };

  const contextSize = (axis:'row'|'col') => {
    const st=useViewer.getState();if(!context || !st.model || !st.sheetName)return;
    const spans=axis==='row'?context.rows:context.cols;
    if(!spans.length)return;
    const start=spans[0].start;const sh=st.model.getSheet(st.sheetName)!;
    st.setSelection({active:axis==='row'?{r:start,c:0}:{r:0,c:start},ranges:spans.map(span=>axis==='row'?{r1:span.start,r2:span.start+span.count-1,c1:0,c2:0}:{r1:0,r2:0,c1:span.start,c2:span.start+span.count-1})});
    setDimKind(axis==='row'?'rowHeight':'colWidth');
    setDimValue(String(axis==='row'?sh.rowHeights.get(start)??30:sh.colWidths.get(start)??(140-5)/7));
    setContext(null);
  };

  const onCellFocused = (e: CellFocusedEvent) => {
    if (e.rowIndex == null || !e.column) return;
    const col = typeof e.column === 'string' ? e.column : e.column.getColId();
    const c = col;
    if (!/^\d+$/.test(c)) return;
    const row = (e.rowPinned==='top' ? e.api.getPinnedTopRow(e.rowIndex) : e.api.getDisplayedRowAtIndex(e.rowIndex))?.data as Row | undefined;
    if (!row) return;
    const r = (row.id as number) - 1;
    const cc = Number(c);
    const st = useViewer.getState();
    if (shiftDown.current && anchor.current && st.selection) {
      const base = anchor.current;
      st.setSelection({ active: { r, c: cc }, ranges: [normRange(base, { r, c: cc })] });
    } else if (!st.selection || st.selection.active.r !== r || st.selection.active.c !== cc) {
      anchor.current = { r, c: cc };
      st.setSelection({ active: { r, c: cc }, ranges: [{ r1: r, c1: cc, r2: r, c2: cc }] });
    }
    grid.current?.refreshCells({ force: true, columns: undefined });
  };

  const onCellClicked = (e: CellClickedEvent) => {
    const colId = typeof e.column === 'string' ? e.column : e.column.getColId();
    const row = e.data as Row | undefined;
    if (!row) return;
    const r = (row.id as number) - 1;
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    if (colId === 'row') {
      const sh = st.model.getSheet(st.sheetName)!;
      select({ r, c: 0 }, [{ r1: r, c1: 0, r2: r, c2: Math.max(0, sh.colCount - 1) }]);
      return;
    }
    if (!/^\d+$/.test(colId)) return;
    const c = Number(colId);
    if (e.event instanceof MouseEvent && (e.event.ctrlKey || e.event.metaKey) && st.selection) {
      select({ r, c }, [...st.selection.ranges, { r1: r, c1: c, r2: r, c2: c }]);
    }
  };

  const onHeaderClick = (e: React.MouseEvent) => {
    if((e.target as HTMLElement).closest('.ag-header-cell-resize,.ag-header-cell-filter-button,.ag-header-cell-menu-button'))return;
    const el = (e.target as HTMLElement).closest('.ag-header-cell');
    if (!el) return;
    const colId = el.getAttribute('col-id');
    if (!colId || !/^\d+$/.test(colId)) return;
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const c = Number(colId);
    const sh = st.model.getSheet(st.sheetName)!;
    select({ r: 0, c }, [{ r1: 0, c1: c, r2: Math.max(0, sh.rowCount - 1), c2: c }]);
  };

  const commitFormulaBar = () => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const { r, c } = st.selection.active;
    const cur = st.model.displayOf(st.sheetName, r, c);
    const curEdit = cur.formula !== undefined ? '=' + cur.formula : cur.text;
    if (fb === curEdit) { setFbEditing(false); return; }
    try {st.exec(editCellsCommand(st.model, st.sheetName, [{ r, c, next: inputToChange(fb,st.model.getRec(st.sheetName,r,c)?.style?.numFmt==='@') }], 'Edit cell'));}catch(e){st.set({error:String(e)});}
    setFbEditing(false);
  };

  // ---------- toolbar actions ----------

  const fmt = (patch: Partial<CellStyle> | null, numFmt?: string, label = 'Format') => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) { setToast('Select cells first'); return; }
    const ranges = st.selection.ranges.length ? st.selection.ranges : [normRange(st.selection.active, st.selection.active)];
    const targets = targetsOf(ranges);
    if (!targets.length) return;
    st.exec(formatCommand(st.model, st.sheetName, targets, patch ?? null, numFmt, label));
  };

  const insertDel = (axis: 'row' | 'col', mode: 'insert' | 'delete') => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName || !st.selection) { setToast('Select a cell first'); return; }
    const at = axis === 'row' ? Math.min(...st.selection.ranges.flatMap((rg) => [rg.r1])) : Math.min(...st.selection.ranges.flatMap((rg) => [rg.c1]));
    const span = axis === 'row'
      ? Math.max(...st.selection.ranges.map((rg) => rg.r2 - rg.r1 + 1))
      : Math.max(...st.selection.ranges.map((rg) => rg.c2 - rg.c1 + 1));
    const count = st.selection.ranges.length > 1 || span > 1 ? span : 1;
    try { st.exec(insertDeleteCommand(st.model, st.sheetName, axis, mode, at, count)); } catch(e) { st.set({error:e instanceof Error?e.message:String(e)}); }
    setToast(`${mode === 'insert' ? 'Inserted' : 'Deleted'} ${count} ${axis}(s)`);
  };

  const hideUnhide = (axis: 'row' | 'col', hide: boolean) => {
    const st = useViewer.getState();
    if (st.busy || !st.model || !st.sheetName || !st.selection) return;
    const sh = st.model.getSheet(st.sheetName)!;
    const ranges = st.selection.ranges.length ? st.selection.ranges : [normRange(st.selection.active, st.selection.active)];
    st.exec(layoutCommand(st.model,st.sheetName,()=>{
    for (const rg of ranges) {
      if (axis === 'row') for (let r = rg.r1; r <= rg.r2; r++) { if (hide) sh.hiddenRows.add(r); else sh.hiddenRows.delete(r); }
      else for (let c = rg.c1; c <= rg.c2; c++) { if (hide) sh.hiddenCols.add(c); else sh.hiddenCols.delete(c); }
    }
    }));
  };

  const unhideAll = () => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const sh = st.model.getSheet(st.sheetName)!;
    st.exec(layoutCommand(st.model,st.sheetName,()=>{sh.hiddenRows.clear();sh.hiddenCols.clear();}));
  };

  const freeze = (rows: number, cols: number) => {
    const st = useViewer.getState();
    if (!st.model || !st.sheetName) return;
    const sh = st.model.getSheet(st.sheetName)!;
    st.exec(layoutCommand(st.model,st.sheetName,()=>{sh.frozenRows=rows;sh.frozenCols=cols;}));
    setFreezeMenu(null);
    setToast(rows || cols ? `Frozen (${rows}R × ${cols}C)` : 'Freeze cleared');
  };

  const freezeAtSelection = () => {
    const st = useViewer.getState();
    if (!st.selection) return;
    freeze(st.selection.active.r, st.selection.active.c);
  };

  const runFind = (all: boolean): Found[] => {
    const st = useViewer.getState();
    if (!st.model || !findText) return [];
    const q = matchCase ? findText : findText.toLowerCase();
    const sheets = findScope === 'sheet' ? [st.sheetName!] : st.model.sheetNames();
    const out: Found[] = [];
    for (const name of sheets) {
      const sh = st.model.getSheet(name);
      if (!sh) continue;
      for (const [k, rec] of sh.cells) {
        const r = Math.floor(k / 16384), c = k % 16384;
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
    if (!list.length) { list = runFind(false); setFound(list); }
    if (!list.length) { setToast('No matches'); return; }
    let i = foundIdx;
    i = i < 0 ? (dir === 1 ? 0 : list.length - 1) : (i + dir + list.length) % list.length;
    setFoundIdx(i);
    gotoFound(list[i]);
  };

  const doReplace = (all: boolean) => {
    const st = useViewer.getState();
    if (!st.model || !findText) return;
    const list = all ? runFind(true) : found.length ? [found[Math.max(0, foundIdx)]] : runFind(false).slice(0, 1);
    if (!list.length) { setToast('No matches'); return; }
    const q = matchCase ? findText : findText.toLowerCase();
    const bySheet = new Map<string, { r: number; c: number; next: { v: string | number | boolean | null; f?: string } }[]>();
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
        const raw=String(rec.v??'');
        nv=matchCase?raw.replaceAll(findText,replaceText):replaceLiteral(raw,findText,replaceText);
        if(nv===raw)continue;
      }
      if (!bySheet.has(f.sheet)) bySheet.set(f.sheet, []);
      bySheet.get(f.sheet)!.push({ r: f.r, c: f.c, next: rec.f===undefined && typeof rec.v==='string' ? {v:nv} : inputToChange(nv) });
    }
    let n = 0;
    const commands=[];
    for (const [name, changes] of bySheet) {
      if (!changes.length) continue;
      commands.push(editCellsCommand(st.model, name, changes, all ? 'Replace all' : 'Replace'));
      n += changes.length;
    }
    if(commands.length)st.exec(compositeCommand(commands,all?'Replace all':'Replace'));
    setToast(n ? `Replaced ${n} cell(s)` : 'No matches');
    if (!all && n) { const l = runFind(false); setFound(l); setFoundIdx(-1); }
  };

  const sheetOp = (op: Parameters<typeof sheetOpCommand>[1]) => {
    const st = useViewer.getState();
    if (!st.model) return;
    if (op.type === 'delete') {
      setDeleteSheet(op.name);
      setSheetMenu(null);
      return;
    }
    if (op.type === 'rename') {
      setRenameSheet(op.oldName);
      setRenameValue(op.oldName);
      setSheetMenu(null);
      return;
    }
    const before = op.type === 'add' || op.type === 'duplicate' ? st.model.sheetNames() : null;
    try {st.exec(sheetOpCommand(st.model, op));}catch(e){st.set({error:e instanceof Error?e.message:String(e)});return;}
    setSheetMenu(null);
    if (before) {
      const created = useViewer.getState().model?.sheetNames().find((n) => !before.includes(n));
      if (created) void useViewer.getState().load(created);
    }
  };

  const confirmRename = () => {
    const st = useViewer.getState();
    if (!st.model || !renameSheet) { setRenameSheet(null); return; }
    const nn = renameValue.trim();
    if (nn && nn !== renameSheet) {try {st.exec(sheetOpCommand(st.model, { type: 'rename', oldName: renameSheet, newName: nn }));}catch(e){st.set({error:String(e)});return;}}
    setRenameSheet(null);
  };

  const confirmDeleteSheet = () => {
    const st = useViewer.getState();
    if (st.model && deleteSheet) {try {st.exec(sheetOpCommand(st.model, { type: 'delete', name: deleteSheet }));}catch(e){st.set({error:String(e)});}}
    setDeleteSheet(null);
  };

  const applyDim = () => {
    const st = useViewer.getState();
    if (st.model && st.sheetName && st.selection && dimKind) {
      const selection=st.selection;
      const sizes=new Map<number,number>();
      const axis=dimKind==='colWidth'?'col':'row';
      const n=axis==='col'?Math.max(5,Math.min(200,Number(dimValue)||12)):Math.max(20,Math.min(600,Number(dimValue)||30));
      for(const range of selection.ranges)for(let i=axis==='col'?range.c1:range.r1;i<=(axis==='col'?range.c2:range.r2);i++)sizes.set(i,n);
      st.exec(resizeCommand(st.model,st.sheetName,axis,sizes));
    }
    setDimKind(null);
  };

  const copyRows = () => {
    const rows = grid.current?.getSelectedNodes().sort((a, b) => (a.rowIndex ?? 0) - (b.rowIndex ?? 0)) ?? [];
    if (!rows.length) { setToast('Select rows using the checkboxes first'); return; }
    const text = rows.map((n) => Array.from({ length: colCount }, (_, i) => {
      const t = (n.data as Row | undefined)?.cells[String(i)]?.text ?? '';
      return /[\t\r\n"]/.test(t) ? '"' + t.replaceAll('"', '""') + '"' : t;
    }).join('\t')).join('\r\n');
    void copyText(text);
  };

  const selLabel = s.selection
    ? (s.selection.ranges.length > 1 ? `${s.selection.ranges.length} ranges` : rangeLabel(s.selection.ranges[0] ?? normRange(s.selection.active, s.selection.active)))
    : (s.selected?.address ?? '—');

  return (
    <Box className="app" onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setDrag(false); }} onDrop={(e) => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) void s.open(undefined, e.dataTransfer.files[0]); }}>
      {typeof window.viewer.chooseSavePath!=='function' && <Alert severity="error">{DESKTOP_UPDATE_MESSAGE}</Alert>}
      <header>
        <Box>
          <Typography variant="h6" fontWeight={700}>▦ Clear Sheet</Typography>
          <Typography variant="caption" color="text.secondary">A lightweight spreadsheet editor</Typography>
        </Box>
        {s.dirty ? <Chip size="small" label="● Unsaved changes" color="warning" variant="outlined" /> : <Chip size="small" label="Saved" variant="outlined" />}
        <Box flex={1} />
        <Button variant="outlined" disabled={s.busy} onClick={() => void s.open()}>Open file</Button>
        <Button variant="contained" disabled={!s.book || s.busy} onClick={() => {grid.current?.stopEditing();void useViewer.getState().save();}}>Save</Button>
        <Button variant="outlined" disabled={!s.book || s.busy} onClick={() => {grid.current?.stopEditing();void useViewer.getState().saveAs();}}>Save As</Button>
      </header>

      {s.error && <Alert severity="error" onClose={() => s.set({ error: '' })}>{s.error}</Alert>}
      {s.busy && <LinearProgress />}

      {s.book ? (
        <>
          <MuiToolbar className="ribbon" variant="dense" disableGutters>
            <Tooltip title="Undo (Ctrl+Z)"><span><IconButton size="small" disabled={!s.canUndo || s.busy} onClick={s.undo} aria-label="Undo">↩</IconButton></span></Tooltip>
            <Tooltip title="Redo (Ctrl+Y)"><span><IconButton size="small" disabled={!s.canRedo || s.busy} onClick={s.redo} aria-label="Redo">↪</IconButton></span></Tooltip>
            <Divider orientation="vertical" flexItem />
            <Tooltip title="Cut (Ctrl+X)"><span><IconButton size="small" disabled={!s.selection} onClick={() => void doCopy(true)} aria-label="Cut">✂</IconButton></span></Tooltip>
            <Tooltip title="Copy (Ctrl+C)"><span><IconButton size="small" disabled={!s.selection} onClick={() => void doCopy(false)} aria-label="Copy">⧉</IconButton></span></Tooltip>
            <Tooltip title="Paste (Ctrl+V)"><span><IconButton size="small" disabled={!s.selection} onClick={() => void doPaste()} aria-label="Paste">📋</IconButton></span></Tooltip>
            <Button size="small" disabled={!s.selected || s.busy} onClick={() => void copyText(s.selected?.text ?? '')}>Copy cell</Button>
            <Button size="small" disabled={s.busy} onClick={copyRows}>Copy rows</Button>
            <Divider orientation="vertical" flexItem />
            <FormControl size="small" sx={{ minWidth: 110 }}><InputLabel id="font-label">Font</InputLabel>
              <Select labelId="font-label" label="Font" value="" displayEmpty onChange={(e) => fmt({ fontName: String(e.target.value) }, undefined, 'Font')} aria-label="Font family">
                {FONT_CHOICES.map((f) => <MenuItem key={f} value={f}>{f}</MenuItem>)}
              </Select></FormControl>
            <FormControl size="small" sx={{ minWidth: 70 }}><InputLabel id="size-label">Size</InputLabel>
              <Select labelId="size-label" label="Size" value="" displayEmpty onChange={(e) => fmt({ fontSize: Number(e.target.value) }, undefined, 'Font size')} aria-label="Font size">
                {SIZE_CHOICES.map((n) => <MenuItem key={n} value={n}>{n}</MenuItem>)}
              </Select></FormControl>
            <Tooltip title="Bold"><span><IconButton size="small" aria-label="Bold" onClick={() => {
              const cur = s.selection && s.model && s.sheetName ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style?.bold : false;
              fmt({ bold: !cur }, undefined, 'Bold');
            }}><b>B</b></IconButton></span></Tooltip>
            <Tooltip title="Italic"><span><IconButton size="small" aria-label="Italic" onClick={() => {
              const cur = s.selection && s.model && s.sheetName ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style?.italic : false;
              fmt({ italic: !cur }, undefined, 'Italic');
            }}><i>I</i></IconButton></span></Tooltip>
            <Tooltip title="Underline"><span><IconButton size="small" aria-label="Underline" onClick={() => {
              const cur = s.selection && s.model && s.sheetName ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style?.underline : false;
              fmt({ underline: !cur }, undefined, 'Underline');
            }}><u>U</u></IconButton></span></Tooltip>
            <Tooltip title="Text color"><input type="color" aria-label="Text color" style={{ width: 28, height: 28, border: 'none', background: 'none' }} onChange={(e) => fmt({ color: e.target.value }, undefined, 'Text color')} /></Tooltip>
            <Tooltip title="Fill color"><input type="color" aria-label="Fill color" style={{ width: 28, height: 28, border: 'none', background: 'none' }} defaultValue="#ffffff" onChange={(e) => fmt({ bg: e.target.value }, undefined, 'Fill color')} /></Tooltip>
            <Select size="small" value="" displayEmpty aria-label="Alignment" onChange={(e) => { const v = String(e.target.value); if (v==='left'||v==='center'||v==='right') fmt({ hAlign: v }, undefined, 'Alignment'); }}>
              <MenuItem value="">Align</MenuItem><MenuItem value="left">Left</MenuItem><MenuItem value="center">Center</MenuItem><MenuItem value="right">Right</MenuItem>
            </Select>
            <Tooltip title="Wrap text"><span><Button size="small" aria-label="Wrap text" onClick={() => {
              const cur = s.selection && s.model && s.sheetName ? s.model.getRec(s.sheetName, s.selection.active.r, s.selection.active.c)?.style?.wrap : false;
              fmt({ wrap: !cur }, undefined, 'Wrap text');
            }}>Wrap</Button></span></Tooltip>
            <Tooltip title="All borders"><span><Button size="small" aria-label="Borders" onClick={() => fmt({ border: { top: true, bottom: true, left: true, right: true } }, undefined, 'Borders')}>▦</Button></span></Tooltip>
            <Tooltip title="Clear formatting"><span><Button size="small" aria-label="Clear formatting" onClick={() => fmt(null, undefined, 'Clear formatting')}>✕Fmt</Button></span></Tooltip>
            <Divider orientation="vertical" flexItem />
            <FormControl size="small" sx={{ minWidth: 130 }}><InputLabel id="num-label">Number</InputLabel>
              <Select labelId="num-label" label="Number" value="" displayEmpty aria-label="Number format" onChange={(e) => { const v = String(e.target.value); if (v) fmt({}, v, 'Number format'); }}>
                {NUM_FORMATS.map((f) => <MenuItem key={f.label} value={f.fmt}>{f.label}</MenuItem>)}
              </Select></FormControl>
            <Divider orientation="vertical" flexItem />
            <Button size="small" aria-label="Insert menu" onClick={(e) => setInsertMenu(e.currentTarget)}>Insert/Delete ▾</Button>
            <Menu anchorEl={insertMenu} open={!!insertMenu} onClose={() => setInsertMenu(null)}>
              <MenuItem onClick={() => { setInsertMenu(null); insertDel('row', 'insert'); }}>Insert row(s)</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); insertDel('row', 'delete'); }}>Delete row(s)</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); insertDel('col', 'insert'); }}>Insert column(s)</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); insertDel('col', 'delete'); }}>Delete column(s)</MenuItem>
              <MenuItem onClick={() => {
                setInsertMenu(null);
                setDimValue('12');
                setDimKind('colWidth');
              }}>Column width…</MenuItem>
              <MenuItem onClick={() => {
                setInsertMenu(null);
                setDimValue('30');
                setDimKind('rowHeight');
              }}>Row height…</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); hideUnhide('row', true); }}>Hide rows</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); hideUnhide('col', true); }}>Hide columns</MenuItem>
              <MenuItem onClick={() => { setInsertMenu(null); unhideAll(); }}>Unhide all</MenuItem>
            </Menu>
            <Button size="small" onClick={() => { grid.current?.setFilterModel(null); grid.current?.applyColumnState({ defaultState: { sort: null } }); setQuery(''); }}>Reset filters</Button>
            <Typography variant="caption" color="text.secondary">Visual styles and freeze: session only</Typography>
            <Button size="small" aria-label="Freeze menu" onClick={(e) => setFreezeMenu(e.currentTarget)}>Freeze ▾</Button>
            <Menu anchorEl={freezeMenu} open={!!freezeMenu} onClose={() => setFreezeMenu(null)}>
              <MenuItem onClick={() => freeze(1, 0)}>Freeze first row</MenuItem>
              <MenuItem onClick={() => freeze(0, 1)}>Freeze first column</MenuItem>
              <MenuItem onClick={freezeAtSelection}>Freeze at selection</MenuItem>
              <MenuItem onClick={() => freeze(0, 0)}>Unfreeze</MenuItem>
            </Menu>
            <Button size="small" onClick={() => setFindOpen(true)}>Find/Replace</Button>
            <Box flex={1} />
            <TextField size="small" placeholder="Search sheet · Ctrl+F" value={query} inputRef={searchRef}
              onChange={(e) => setQuery(e.target.value)} slotProps={{ htmlInput: { 'aria-label': 'Search sheet' } }} />
            <Select size="small" value={s.zoom} inputProps={{ 'aria-label': 'Zoom' }} onChange={(e) => s.set({ zoom: Number(e.target.value) })}>
              {[75, 90, 100, 110, 125, 150, 175, 200].map((n) => <MenuItem key={n} value={n}>{n}%</MenuItem>)}
            </Select>
          </MuiToolbar>

          <section className="toolbar">
            <Typography fontWeight={600} noWrap sx={{ maxWidth: 300 }} title={s.book.name}>{s.book.name}</Typography>
            {s.filePath && <Typography variant="caption" color="text.secondary" noWrap sx={{ maxWidth: 420 }} title={s.filePath}>{s.filePath}</Typography>}
          </section>

          <section className="formula">
            <span className="address" title="Active cell / range">{selLabel}</span>
            <span className="fx">fx</span>
            <input aria-label="Formula bar" value={fb} placeholder="Select a cell to view or edit its value or formula"
              onFocus={() => setFbEditing(true)}
              onChange={(e) => { setFb(e.target.value); setFbEditing(true); }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') { e.preventDefault(); commitFormulaBar(); (e.target as HTMLInputElement).blur(); if(s.selection)focusCell(s.selection.active.r,s.selection.active.c); }
                if (e.key === 'Escape') { setFb(s.selected?.formula !== undefined ? '=' + s.selected.formula : s.selected?.text ?? ''); setFbEditing(false); (e.target as HTMLInputElement).blur(); }
                e.stopPropagation();
              }}
              onBlur={() => setFbEditing(false)} />
          </section>

          {s.model?.requiresCopy && <Alert severity="info">Save creates an edited copy to preserve the original Excel file. Fonts, fills, borders, freeze panes, charts and other advanced features may not persist in that copy.</Alert>}
          {s.model && s.sheetName && s.model.getSheet(s.sheetName)!.rowCount > 200000 && <Alert severity="warning">Grid shows the first 200,000 rows; the complete model is retained for saving.</Alert>}
          {s.sheet?.truncated && <Alert severity="warning">Preview limited to {s.sheet.rows.length.toLocaleString()} rows and {s.sheet.columns} columns (up to 1 million cell positions). This sheet has {s.sheet.totalRows.toLocaleString()} rows. Search, sort and filters cover this preview only.</Alert>}

          <main className="grid" aria-busy={s.busy} onClick={onHeaderClick} onContextMenu={openHeaderMenu}>
            {sheetName && (
              <AgGridReact<Row>
                key={s.book.name + sheetName}
                theme={gridTheme}
                rowData={pinnedTop?.length ? s.rows.filter(row=>!pinnedTop.some(p=>p.id===row.id)) : s.rows}
                columnDefs={columns}
                defaultColDef={{ editable: false }}
                getRowId={(p) => String(p.data.id)}
                quickFilterText={s.search}
                pinnedTopRowData={pinnedTop as Row[] | undefined}
                getRowHeight={getRowHeight}
                stopEditingWhenCellsLoseFocus
                rowSelection={{mode:"multiRow",enableClickSelection:false}}
                onColumnResized={onColumnResized}
                onCellContextMenu={openCellMenu}
                preventDefaultOnContextMenu
                onGridReady={(e) => { grid.current = e.api; }}
                onCellFocused={onCellFocused}
                onCellClicked={onCellClicked}
                overlayNoRowsTemplate="No matching rows"
              />
            )}
            {s.busy && <div className="loading">Reading spreadsheet…</div>}
          </main>

          <Menu open={!!context} onClose={()=>setContext(null)} anchorReference="anchorPosition" anchorPosition={context?{top:context.y,left:context.x}:undefined}>
            {!!context?.rows.length && <MenuItem onClick={()=>contextOperation('row','insert')}>Insert row(s) above</MenuItem>}
            {!!context?.rows.length && <MenuItem onClick={()=>contextOperation('row','delete')}>Delete row(s)</MenuItem>}
            {!!context?.rows.length && <MenuItem onClick={()=>contextSize('row')}>Row height…</MenuItem>}
            {!!context?.rows.length && <Divider/>}
            <MenuItem onClick={()=>contextOperation('col','insert')}>Insert column(s) before</MenuItem>
            <MenuItem onClick={()=>contextOperation('col','delete')}>Delete column(s)</MenuItem>
            <MenuItem onClick={()=>contextSize('col')}>Column width…</MenuItem>
          </Menu>

          <Box display="flex" alignItems="center">
            <Tabs value={sheetName ?? false} onChange={(_, name) => void s.load(name)} variant="scrollable" scrollButtons="auto" sx={{ flex: 1 }}>
              {(s.book.sheets).map((name) => <Tab key={name} label={name} value={name} disabled={s.busy} />)}
            </Tabs>
            <Button size="small" aria-label="Add sheet" onClick={() => sheetOp({ type: 'add' })}>+ Sheet</Button>
            <Button size="small" aria-label="Sheet menu" disabled={!sheetName} onClick={(e) => setSheetMenu(e.currentTarget)}>Sheet ▾</Button>
            <Menu anchorEl={sheetMenu} open={!!sheetMenu} onClose={() => setSheetMenu(null)}>
              <MenuItem onClick={() => sheetName && sheetOp({ type: 'rename', oldName: sheetName, newName: '' })}>Rename…</MenuItem>
              <MenuItem onClick={() => sheetName && sheetOp({ type: 'duplicate', name: sheetName })}>Duplicate</MenuItem>
              <MenuItem onClick={() => {
                const st = useViewer.getState();
                if (sheetName && st.model) {
                  const i = st.model.sheetNames().indexOf(sheetName);
                  st.exec(sheetOpCommand(st.model, { type: 'move', name: sheetName, from: i, to: Math.max(0, i - 1) }));
                }
                setSheetMenu(null);
              }}>Move left</MenuItem>
              <MenuItem onClick={() => {
                const st = useViewer.getState();
                if (sheetName && st.model) {
                  const i = st.model.sheetNames().indexOf(sheetName);
                  st.exec(sheetOpCommand(st.model, { type: 'move', name: sheetName, from: i, to: i + 1 }));
                }
                setSheetMenu(null);
              }}>Move right</MenuItem>
              <MenuItem onClick={() => sheetName && sheetOp({ type: 'delete', name: sheetName })}>Delete…</MenuItem>
            </Menu>
          </Box>

          <footer>
            <span>{sheetName && model ? `${model.getSheet(sheetName)?.rowCount.toLocaleString()} rows · ${model.getSheet(sheetName)?.colCount} columns` : 'Choose a sheet'}</span>
            <span>{s.dirty ? '● Unsaved changes (Ctrl+S to save)' : 'All changes saved'} · {selLabel} · Double-click / F2 / type to edit</span>
          </footer>
        </>
      ) : (
        <main className="welcome">
          <div className="file-icon">▦</div>
          <Typography variant="h4" fontWeight={650}>Your spreadsheets, simply edited.</Typography>
          <Typography color="text.secondary">Drop a file here or open one from your computer.</Typography>
          <Typography variant="caption" color="text.secondary">XLSX · XLS · XLSM · XLSB · CSV</Typography>
          <Button variant="contained" size="large" disabled={s.busy} onClick={() => void s.open()}>Browse files</Button>
          <section className="recent">
            <Box display="flex" justifyContent="space-between" alignItems="center">
              <Typography fontWeight={600}>Recent files</Typography>
              <Button size="small" onClick={async () => { const r = await window.viewer.clearRecent(); if (r.ok) void s.refresh(); else s.set({ error: r.error }); }}>Clear history</Button>
            </Box>
            {s.recents.length ? s.recents.map((path) => <Button className="recent-file" key={path} disabled={s.busy} onClick={() => void s.open(path)} title={path}>{path}</Button>) : <Typography color="text.secondary" variant="body2">Opened files will appear here.</Typography>}
          </section>
        </main>
      )}

      <Dialog open={findOpen} onClose={() => setFindOpen(false)} maxWidth="sm" fullWidth>
        <DialogTitle>Find / Replace</DialogTitle>
        <DialogContent sx={{ display: 'flex', flexDirection: 'column', gap: 2, pt: 1 }}>
          <TextField label="Find" value={findText} onChange={(e) => { setFindText(e.target.value); setFound([]); setFoundIdx(-1); }} autoFocus
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); findNav(1); } }} />
          <TextField label="Replace with" value={replaceText} onChange={(e) => setReplaceText(e.target.value)} />
          <Box display="flex" gap={2}>
            <FormControl size="small" sx={{ minWidth: 140 }}>
              <InputLabel>Scope</InputLabel>
              <Select value={findScope} label="Scope" onChange={(e) => setFindScope(e.target.value==='book'?'book':'sheet')}>
                <MenuItem value="sheet">Current sheet</MenuItem><MenuItem value="book">Workbook</MenuItem>
              </Select>
            </FormControl>
            <FormControlLabel control={<Checkbox checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />} label="Match case" />
          </Box>
          {found.length > 0 && <Typography variant="caption">{foundIdx + 1} / {found.length} matches</Typography>}
          {found.length > 0 && (
            <List dense sx={{ maxHeight: 200, overflow: 'auto', border: '1px solid #eee' }}>
              {found.slice(0, 500).map((f, i) => (
                <ListItemButton key={`${f.sheet}:${f.r}:${f.c}`} selected={i === foundIdx} onClick={() => { setFoundIdx(i); gotoFound(f); }}>
                  <Typography variant="body2">{f.sheet}!{colToLetter(f.c)}{f.r + 1}: {f.text.slice(0, 80)}</Typography>
                </ListItemButton>
              ))}
            </List>
          )}
        </DialogContent>
        <DialogActions>
          <Button onClick={() => findNav(-1)}>Find prev</Button>
          <Button onClick={() => findNav(1)}>Find next</Button>
          <Button onClick={() => { const l = runFind(true); setFound(l); setFoundIdx(l.length ? 0 : -1); if (!l.length) setToast('No matches'); }}>Find all</Button>
          <Button onClick={() => doReplace(false)}>Replace</Button>
          <Button onClick={() => doReplace(true)}>Replace all</Button>
          <Button onClick={() => setFindOpen(false)}>Close</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={renameSheet !== null} onClose={() => setRenameSheet(null)} maxWidth="xs" fullWidth>
        <DialogTitle>Rename worksheet</DialogTitle>
        <DialogContent>
          <TextField label="Sheet name" value={renameValue} onChange={(e) => setRenameValue(e.target.value)} fullWidth autoFocus
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); confirmRename(); } }} />
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setRenameSheet(null)}>Cancel</Button>
          <Button variant="contained" onClick={confirmRename}>Rename</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={deleteSheet !== null} onClose={() => setDeleteSheet(null)} maxWidth="xs" fullWidth>
        <DialogTitle>Delete worksheet?</DialogTitle>
        <DialogContent><Typography>Delete worksheet “{deleteSheet}”? You can undo this while the workbook remains open.</Typography></DialogContent>
        <DialogActions>
          <Button onClick={() => setDeleteSheet(null)}>Cancel</Button>
          <Button variant="contained" color="error" onClick={confirmDeleteSheet}>Delete</Button>
        </DialogActions>
      </Dialog>

      <Dialog open={dimKind !== null} onClose={() => setDimKind(null)} maxWidth="xs" fullWidth>
        <DialogTitle>{dimKind === 'colWidth' ? 'Column width (characters)' : 'Row height (pixels)'}</DialogTitle>
        <DialogContent>
          <TextField label={dimKind === 'colWidth' ? 'Width' : 'Height'} value={dimValue} onChange={(e) => setDimValue(e.target.value)} fullWidth autoFocus inputProps={{ inputMode: 'numeric' }}
            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); applyDim(); } }} />
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setDimKind(null)}>Cancel</Button>
          <Button variant="contained" onClick={applyDim}>Apply</Button>
        </DialogActions>
      </Dialog>

      {drag && <div className="drop-overlay">Drop spreadsheet to open</div>}
      <Snackbar open={!!toast} autoHideDuration={2200} onClose={() => setToast('')} message={toast} />
    </Box>
  );
}

function replaceLiteral(src: string, find: string, repl: string): string {
  if (!find) return src;
  return src.toLowerCase().split(find.toLowerCase()).join(repl);
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><ThemeProvider theme={theme}><CssBaseline /><App /></ThemeProvider></React.StrictMode>);
