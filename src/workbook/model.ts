import * as XLSX from 'xlsx';
import type { CellStyle } from '../../electron/types';
import { colToLetter } from './cellRef';
import { formatValue, styleToXlsx, xlsxToStyle, dateToExcelSerial } from './format';
import { FormulaEngine, isErr, shiftFormula, renameFormulaSheet, type EvalVal, type FormulaSource } from './formulas';

export type RawVal = string | number | boolean | null;

export interface CellRec {
  v: RawVal;
  original?: XLSX.CellObject;
  edited?: boolean;
  f?: string;
  style?: CellStyle;
  /** last computed value for formula cells */
  calc?: EvalVal;
}

export interface SheetData {
  name: string;
  original?: XLSX.WorkSheet;
  metadata?: NonNullable<XLSX.WorkBook["Workbook"]>["Sheets"] extends (infer T)[] | undefined ? T : never;
  cells: Map<number, CellRec>; // key = r*16384+c
  rowCount: number;
  colCount: number;
  colWidths: Map<number, number>; // wch
  rowHeights: Map<number, number>; // px
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  frozenRows: number;
  frozenCols: number;
}

export const keyOf = (r: number, c: number) => r * 16384 + c;
export const rcOf = (k: number) => ({ r: Math.floor(k / 16384), c: k % 16384 });

function cellAddr(r: number, c: number): string {
  return `${colToLetter(c)}${r + 1}`;
}

export interface SheetSnapshot {
  name: string;
  original?: XLSX.WorkSheet;
  metadata?: NonNullable<XLSX.WorkBook["Workbook"]>["Sheets"] extends (infer T)[] | undefined ? T : never;
  cells: Map<number, CellRec>;
  rowCount: number; colCount: number;
  colWidths: Map<number, number>; rowHeights: Map<number, number>;
  hiddenRows: Set<number>; hiddenCols: Set<number>;
  frozenRows: number; frozenCols: number;
}

function cloneRec(c: CellRec): CellRec {
  return { ...c, v: c.v, ...(c.f !== undefined ? { f: c.f } : {}), ...(c.style ? { style: { ...c.style, ...(c.style.border ? { border: { ...c.style.border } } : {}) } } : {}), ...(c.calc !== undefined ? { calc: c.calc } : {}) };
}

/**
 * Internal workbook representation (sparse). AG Grid is a view over this model.
 * Keeps the original SheetJS workbook for save-time preservation of
 * unsupported features (charts, macros, pivots, etc.).
 */
export class WorkbookModel implements FormulaSource {
  sheets: SheetData[] = [];
  wb: XLSX.WorkBook | null = null;
  fileType: string = 'xlsx';
  dirty = false;
  requiresCopy = false;
  sourcePath: string | null = null;

  engine: FormulaEngine;

  constructor() {
    this.engine = new FormulaEngine(this);
  }

  sheetNames(): string[] { return this.sheets.map((s) => s.name); }

  getSheet(name: string): SheetData | undefined {
    return this.sheets.find((s) => s.name === name);
  }

  resolveSheet(name: string | undefined, from: string): string | null {
    if (!name) return this.getSheet(from) ? from : null;
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower)?.name ?? null;
  }

  rawValue(sheet: string, r: number, c: number): EvalVal {
    const sh = this.getSheet(sheet);
    const rec = sh?.cells.get(keyOf(r, c));
    if (!rec) return null;
    if (!rec.edited && rec.original?.t==='e') {const errors:Record<number,string>={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A'};return {err:errors[Number(rec.original.v)]??'#VALUE!'};}
    if (rec.f !== undefined) {
      const cached = this.engine.cachedValue(sheet, r, c);
      if (cached !== undefined) {
        if (isErr(cached)) return null;
        return (cached ?? null) as RawVal;
      }
      return rec.v;
    }
    return rec.v;
  }

  nonEmptyCount(sheet: string, r1: number, c1: number, r2: number, c2: number): number {
    const sh = this.getSheet(sheet);
    if (!sh) return 0;
    let n = 0;
    for (const [k, rec] of sh.cells) {
      const { r, c } = rcOf(k);
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) {
        if (rec.v !== null && rec.v !== '' || rec.f !== undefined) n++;
      }
    }
    // formula cells tracked only in sh.cells too (we store rec for formulas), so covered.
    return n;
  }

  sparseCells(sheet: string, r1: number, c1: number, r2: number, c2: number): { r: number; c: number }[] {
    const sh = this.getSheet(sheet);
    if (!sh) return [];
    const out: { r: number; c: number }[] = [];
    for (const k of sh.cells.keys()) {
      const { r, c } = rcOf(k);
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) out.push({ r, c });
    }
    return out;
  }

  // ---------- loading ----------

  loadFromBytes(bytes: Uint8Array | Buffer, fileType: string): void {
    this.fileType = fileType.toLowerCase();
    const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellStyles: true, cellNF: true, cellDates: false, bookVBA: true, bookFiles: true, sheetStubs: true, sheetRows: 0 });
    this.wb = wb;
    // CE cannot round-trip arbitrary OOXML parts (drawings, pivots, rich styles).
    // Keep the input untouched: Excel inputs must first be saved as a separate copy.
    this.requiresCopy = this.fileType !== 'csv';
    this.sheets = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const sh: SheetData = {
        name, original: ws, metadata: wb.Workbook?.Sheets?.find(s => s.name === name), cells: new Map(), rowCount: 0, colCount: 0,
        colWidths: new Map(), rowHeights: new Map(),
        hiddenRows: new Set(), hiddenCols: new Set(), frozenRows: 0, frozenCols: 0,
      };
      const ref = ws['!ref'];
      if (ref) {
        const range = XLSX.utils.decode_range(ref);
        sh.rowCount = range.e.r + 1;
        sh.colCount = range.e.c + 1;
      }
      const cols = ws['!cols'] as { width?: number; wch?: number; hidden?: boolean }[] | undefined;
      if (cols) cols.forEach((col, i) => {
        // SheetJS infers MDW globally, which can give different wch/wpx on a
        // later open. Use the stored Excel width with our fixed 7px grid metric.
        if (col?.width !== undefined) sh.colWidths.set(i, Math.max(1, col.width - 5 / 7));
        else if (col?.wch) sh.colWidths.set(i, col.wch);
        if (col?.hidden) sh.hiddenCols.add(i);
      });
      const rows = ws['!rows'] as { hpx?: number; hidden?: boolean }[] | undefined;
      if (rows) rows.forEach((row, i) => {
        if (row?.hpx) sh.rowHeights.set(i, row.hpx);
        if (row?.hidden) sh.hiddenRows.add(i);
      });
      const freeze = ws['!freeze'] as string | undefined;
      if (freeze) {
        const m = /([A-Z]+)(\d+)/.exec(freeze);
        if (m) {
          let c = 0;
          for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
          sh.frozenCols = c - 1; // freeze pane at cell => cols left of it
          sh.frozenRows = parseInt(m[2], 10) - 1;
        }
      }
      for (const addr of Object.keys(ws)) {
        if (addr.startsWith('!')) continue;
        let rc: { r: number; c: number };
        try { rc = XLSX.utils.decode_cell(addr); } catch { continue; }
        const cell = ws[addr] as { v?: unknown; f?: string; w?: string; s?: unknown; z?: string; t?: string };
        const k = keyOf(rc.r, rc.c);
        let v: RawVal = null;
        if (cell.t === 'b') v = cell.v === true || cell.v === 1 ? true : cell.v === false || cell.v === 0 ? false : null;
        else if (cell.t === 'n') v = typeof cell.v === 'number' ? cell.v : Number(cell.v ?? 0);
        else if (cell.t === 'e') v = null;
        else if (cell.t === 'd' && cell.v instanceof Date) v = dateToExcelSerial(cell.v);
        else if (cell.v === undefined || cell.v === null) v = null;
        else if (typeof cell.v === 'string' || typeof cell.v === 'number' || typeof cell.v === 'boolean') v = cell.v;
        else v = String(cell.v);
        const style = xlsxToStyle(cell);
        const rec: CellRec = { v, original: ws[addr], ...(cell.f ? { f: cell.f } : {}), ...(style ? { style } : {}) };
        sh.cells.set(k, rec);
        if (rc.r + 1 > sh.rowCount) sh.rowCount = rc.r + 1;
        if (rc.c + 1 > sh.colCount) sh.colCount = rc.c + 1;
      }
      if (sh.rowCount === 0) sh.rowCount = 50;
      if (sh.colCount === 0) sh.colCount = 10;
      this.sheets.push(sh);
    }
    if (!this.sheets.length) {
      this.sheets.push({
        name: 'Sheet1', cells: new Map(), rowCount: 100, colCount: 10,
        colWidths: new Map(), rowHeights: new Map(), hiddenRows: new Set(), hiddenCols: new Set(), frozenRows: 0, frozenCols: 0,
      });
    }
    this.rebuildFormulas();
    this.dirty = false;
  }

  /** Build preview model from worker rows (fast path before bytes arrive). */
  loadFromPreview(name: string, rows: { id: number; cells: Record<string, { v: RawVal; text: string; formula?: string }> }[], columns: number, totalRows: number): SheetData {
    let sh = this.getSheet(name);
    if (!sh) {
      sh = {
        name, cells: new Map(), rowCount: Math.max(totalRows, 50), colCount: Math.max(columns, 10),
        colWidths: new Map(), rowHeights: new Map(), hiddenRows: new Set(), hiddenCols: new Set(), frozenRows: 0, frozenCols: 0,
      };
      this.sheets.push(sh);
    }
    for (const row of rows) {
      const r = row.id - 1;
      for (const [cs, cell] of Object.entries(row.cells)) {
        const c = Number(cs);
        const k = keyOf(r, c);
        const existing = sh.cells.get(k);
        if (existing && (existing.f !== undefined || existing.style)) continue; // full load wins
        if (cell.v !== null || cell.formula) {
          sh.cells.set(k, { v: cell.v, ...(cell.formula ? { f: cell.formula } : {}) });
        }
      }
      if (r + 1 > sh.rowCount) sh.rowCount = r + 1;
    }
    if (columns > sh.colCount) sh.colCount = columns;
    this.rebuildFormulas();
    return sh;
  }

  rebuildFormulas(): void {
    // clear engine state by re-creating (cheap: only formula cells re-indexed)
    this.engine = new FormulaEngine(this);
    for (const sh of this.sheets) {
      for (const [k, rec] of sh.cells) {
        if (rec.f !== undefined) {
          const { r, c } = rcOf(k);
          rec.calc = this.engine.setFormula(sh.name, r, c, rec.f);
        }
      }
    }
    this.engine.recalcAll();
    // refresh cached values
    for (const sh of this.sheets) {
      for (const [k, rec] of sh.cells) {
        if (rec.f !== undefined) {
          const { r, c } = rcOf(k);
          rec.calc = this.engine.cachedValue(sh.name, r, c);
        }
      }
    }
  }

  refreshCalc(sheet: string, r: number, c: number, val: EvalVal): void {
    const sh = this.getSheet(sheet);
    const rec = sh?.cells.get(keyOf(r, c));
    if (rec && rec.f !== undefined) rec.calc = val;
  }

  // ---------- display ----------

  displayOf(sheet: string, r: number, c: number): { v: RawVal; text: string; formula?: string } {
    const sh = this.getSheet(sheet);
    const rec = sh?.cells.get(keyOf(r, c));
    if (!rec) return { v: null, text: '' };
    if (rec.f !== undefined) {
      const calc = rec.calc ?? this.engine.cachedValue(sheet, r, c) ?? null;
      const v: RawVal = isErr(calc) ? null : ((calc ?? null) as RawVal);
      const text = isErr(calc) ? calc.err : formatValue(v, rec.style?.numFmt);
      return { v, text, formula: rec.f };
    }
    return { v: rec.v, text: !rec.edited && rec.original?.t === 'e' ? rec.original.w ?? '#VALUE!' : formatValue(rec.v, rec.style?.numFmt) };
  }

  cellTextForCopy(sheet: string, r: number, c: number, withFormulas: boolean): string {
    const sh = this.getSheet(sheet);
    const rec = sh?.cells.get(keyOf(r, c));
    if (!rec) return '';
    if (rec.f !== undefined && withFormulas) return '=' + rec.f;
    if (rec.f !== undefined) {
      const calc = rec.calc ?? null;
      if (isErr(calc)) return calc.err;
      return formatValue((calc ?? null) as RawVal, rec.style?.numFmt);
    }
    if (typeof rec.v === 'boolean') return rec.v ? 'TRUE' : 'FALSE';
    if (rec.v === null) return '';
    return formatValue(rec.v, rec.style?.numFmt);
  }

  // ---------- mutation primitives (no undo here; commands wrap these) ----------

  ensureSize(sh: SheetData, r: number, c: number): void {
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= 1048576 || c >= 16384) throw new Error('Cell exceeds Excel row/column limits.');
    if (r >= sh.rowCount) sh.rowCount = r + 1;
    if (c >= sh.colCount) sh.colCount = c + 1;
    if (sh.rowCount > 1048576) sh.rowCount = 1048576;
    if (sh.colCount > 16384) sh.colCount = 16384;
  }

  getRec(sheet: string, r: number, c: number): CellRec | undefined {
    return this.getSheet(sheet)?.cells.get(keyOf(r, c));
  }

  /** Set raw cell (value+formula). Returns previous rec (cloned) for undo. */
  setCell(sheet: string, r: number, c: number, v: RawVal, f?: string): CellRec | undefined {
    const sh = this.getSheet(sheet);
    if (!sh) return undefined;
    this.ensureSize(sh, r, c);
    const k = keyOf(r, c);
    const prev = sh.cells.get(k);
    const prevClone = prev ? cloneRec(prev) : undefined;
    if (prev?.f !== undefined) this.engine.removeFormula(sheet, r, c);
    if (v === null && f === undefined) {
      // keep style-only cell if a style exists
      if (prev?.style) sh.cells.set(k, { ...prev, v: null, f: undefined, calc: undefined, edited: true, style: prev.style });
      else sh.cells.delete(k);
    } else {
      const rec: CellRec = { original: prev?.original, edited: true, v, ...(f !== undefined ? { f } : {}), ...(prev?.style ? { style: prev.style } : {}) };
      if (f !== undefined) rec.calc = this.engine.setFormula(sheet, r, c, f);
      sh.cells.set(k, rec);
    }
    this.dirty = true;
    return prevClone;
  }

  /** Restore a rec snapshot (undo). */
  restoreCell(sheet: string, r: number, c: number, prev: CellRec | undefined): void {
    const sh = this.getSheet(sheet);
    if (!sh) return;
    const k = keyOf(r, c);
    const cur = sh.cells.get(k);
    if (cur?.f !== undefined) this.engine.removeFormula(sheet, r, c);
    if (!prev) sh.cells.delete(k);
    else {
      const rec = cloneRec(prev);
      if (rec.f !== undefined) rec.calc = this.engine.setFormula(sheet, r, c, rec.f);
      sh.cells.set(k, rec);
    }
    this.dirty = true;
  }

  setStyle(sheet: string, r: number, c: number, patch: Partial<CellStyle> | null): CellStyle | undefined {
    const sh = this.getSheet(sheet);
    if (!sh) return undefined;
    this.ensureSize(sh, r, c);
    const k = keyOf(r, c);
    const prev = sh.cells.get(k)?.style ? { ...sh.cells.get(k)!.style! } : undefined;
    if (patch === null) {
      const rec = sh.cells.get(k);
      if (rec) {
        delete rec.style;
        if (rec.v === null && rec.f === undefined) sh.cells.delete(k);
      }
    } else {
      let rec = sh.cells.get(k);
      if (!rec) { rec = { v: null }; sh.cells.set(k, rec); }
      rec.style = { ...(rec.style ?? {}), ...patch };
      // drop undefined keys
      for (const kk of Object.keys(rec.style) as (keyof CellStyle)[]) {
        if (rec.style[kk] === undefined) delete rec.style[kk];
      }
      if (!Object.keys(rec.style).length) delete rec.style;
    }
    this.dirty = true;
    return prev;
  }

  /** Notify engine that raw cells changed; returns affected formula addrs with new values. */
  afterCellsChanged(changes: { sheet: string; r: number; c: number }[]): { sheet: string; r: number; c: number; val: EvalVal }[] {
    const updated = this.engine.recalc(changes);
    const out: { sheet: string; r: number; c: number; val: EvalVal }[] = [];
    for (const [k, val] of updated) {
      const idx = k.indexOf('\0');
      const sheet = k.slice(0, idx);
      const [rs, cs] = k.slice(idx + 1).split(':');
      const r = Number(rs), c = Number(cs);
      this.refreshCalc(sheet, r, c, val);
      out.push({ sheet, r, c, val });
    }
    return out;
  }

  // ---------- rows / columns ----------

  insertRows(sheet: string, at: number, count: number): void {
    const sh = this.getSheet(sheet);
    if (!sh || count <= 0) return;
    const moved = new Map<number, CellRec>();
    for (const [k, rec] of sh.cells) {
      const { r, c } = rcOf(k);
      if (r >= at) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        moved.set(keyOf(r + count, c), rec);
      }
    }
    for (const [k] of [...sh.cells]) {
      if (rcOf(k).r >= at) sh.cells.delete(k);
    }
    for (const [k, rec2] of moved) sh.cells.set(k, rec2);
    // shift heights/hidden
    const nh = new Map<number, number>();
    for (const [i, h] of sh.rowHeights) nh.set(i >= at ? i + count : i, h);
    sh.rowHeights = nh;
    const hid = new Set<number>();
    for (const i of sh.hiddenRows) hid.add(i >= at ? i + count : i);
    sh.hiddenRows = hid;
    sh.rowCount += count;
    this.shiftDimensions(sh, 'row', 'insert', at, count);
    this.shiftAllFormulas(sheet, 'row', at, count);
    this.dirty = true;
  }

  deleteRows(sheet: string, at: number, count: number): Map<number, CellRec> {
    const sh = this.getSheet(sheet);
    const removed = new Map<number, CellRec>();
    if (!sh || count <= 0) return removed;
    for (const [k, rec] of [...sh.cells]) {
      const { r, c } = rcOf(k);
      if (r >= at && r < at + count) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        removed.set(k, cloneRec(rec));
        sh.cells.delete(k);
      }
    }
    const moved: [number, CellRec][] = [];
    for (const [k, rec] of [...sh.cells]) {
      const { r, c } = rcOf(k);
      if (r >= at + count) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        sh.cells.delete(k);
        moved.push([keyOf(r - count, c), rec]);
      }
    }
    for (const [k, rec] of moved) sh.cells.set(k, rec);
    const nh = new Map<number, number>();
    for (const [i, h] of sh.rowHeights) {
      if (i >= at && i < at + count) continue;
      nh.set(i >= at + count ? i - count : i, h);
    }
    sh.rowHeights = nh;
    const hid = new Set<number>();
    for (const i of sh.hiddenRows) {
      if (i >= at && i < at + count) continue;
      hid.add(i >= at + count ? i - count : i);
    }
    sh.hiddenRows = hid;
    sh.rowCount = Math.max(1, sh.rowCount - count);
    this.shiftDimensions(sh, 'row', 'delete', at, count);
    this.shiftAllFormulas(sheet, 'row', at + count, -count, at);
    this.dirty = true;
    return removed;
  }

  insertCols(sheet: string, at: number, count: number): void {
    const sh = this.getSheet(sheet);
    if (!sh || count <= 0) return;
    const moved = new Map<number, CellRec>();
    for (const [k, rec] of sh.cells) {
      const { r, c } = rcOf(k);
      if (c >= at) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        moved.set(keyOf(r, c + count), rec);
      }
    }
    for (const [k] of [...sh.cells]) {
      if (rcOf(k).c >= at) sh.cells.delete(k);
    }
    for (const [k, rec2] of moved) sh.cells.set(k, rec2);
    const nw = new Map<number, number>();
    for (const [i, w] of sh.colWidths) nw.set(i >= at ? i + count : i, w);
    sh.colWidths = nw;
    const hid = new Set<number>();
    for (const i of sh.hiddenCols) hid.add(i >= at ? i + count : i);
    sh.hiddenCols = hid;
    sh.colCount += count;
    this.shiftDimensions(sh, 'col', 'insert', at, count);
    this.shiftAllFormulas(sheet, 'col', at, count);
    this.dirty = true;
  }

  deleteCols(sheet: string, at: number, count: number): Map<number, CellRec> {
    const sh = this.getSheet(sheet);
    const removed = new Map<number, CellRec>();
    if (!sh || count <= 0) return removed;
    for (const [k, rec] of [...sh.cells]) {
      const { r, c } = rcOf(k);
      if (c >= at && c < at + count) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        removed.set(k, cloneRec(rec));
        sh.cells.delete(k);
      }
    }
    const moved: [number, CellRec][] = [];
    for (const [k, rec] of [...sh.cells]) {
      const { r, c } = rcOf(k);
      if (c >= at + count) {
        if (rec.f !== undefined) this.engine.removeFormula(sheet, r, c);
        sh.cells.delete(k);
        moved.push([keyOf(r, c - count), rec]);
      }
    }
    for (const [k, rec] of moved) sh.cells.set(k, rec);
    const nw = new Map<number, number>();
    for (const [i, w] of sh.colWidths) {
      if (i >= at && i < at + count) continue;
      nw.set(i >= at + count ? i - count : i, w);
    }
    sh.colWidths = nw;
    const hid = new Set<number>();
    for (const i of sh.hiddenCols) {
      if (i >= at && i < at + count) continue;
      hid.add(i >= at + count ? i - count : i);
    }
    sh.hiddenCols = hid;
    sh.colCount = Math.max(1, sh.colCount - count);
    this.shiftDimensions(sh, 'col', 'delete', at, count);
    this.shiftAllFormulas(sheet, 'col', at + count, -count, at);
    this.dirty = true;
    return removed;
  }

  private shiftDimensions(sh: SheetData, axis: 'row'|'col', mode: 'insert'|'delete', at: number, count: number): void {
    if(sh.original) {
      const props: XLSX.WorkSheet={};
      for(const [k,v] of Object.entries(sh.original)) if(k.startsWith('!'))props[k]=v;
      const field=axis==='row'?'!rows':'!cols';
      const entries=sh.original[field];
      if(entries) {
        const moved: typeof entries=[];
        entries.forEach((v,i)=>{if(mode==='delete'&&i>=at&&i<at+count)return;const to=i>=at?(mode==='insert'?i+count:i-count):i;moved[to]=v;});
        props[field]=moved;
      }
      sh.original=props;
    }
    const field=axis==='row'?'frozenRows':'frozenCols';
    if(at<sh[field]) sh[field]=mode==='insert'?sh[field]+count:Math.max(at,sh[field]-count);
  }

  /** Rewrite formula expressions + re-index engine after row/col insert/delete.
   *  Cell positions in `sh.cells` are already moved by the caller; this only
   *  fixes expressions (pointing into the new layout) and re-registers them.
   *  `delFrom` (>=0 on delete) marks the removed span for #REF! handling. */
  private shiftAllFormulas(sheet: string, axis: 'row' | 'col', at: number, delta: number, delFrom = -1): void {
    // collect all formula cells across workbook (cross-sheet refs may point here)
    const updates: { sheet: string; r: number; c: number; rec: CellRec; expr: string }[] = [];
    for (const sh of this.sheets) {
      for (const [k, rec] of sh.cells) {
        if (rec.f === undefined) continue;
        const { r, c } = rcOf(k);
        this.engine.removeFormula(sh.name, r, c);
        const expr = shiftFormula(rec.f, sheet, axis, at, delta, delFrom, sh.name);
        updates.push({ sheet: sh.name, r, c, rec, expr });
      }
    }
    for (const u of updates) {
      u.rec.f = u.expr;
      u.rec.calc = this.engine.setFormula(u.sheet, u.r, u.c, u.expr);
    }
    this.engine.recalcAll();
    for (const sh of this.sheets) {
      for (const [k, rec] of sh.cells) {
        if (rec.f !== undefined) {
          const { r, c } = rcOf(k);
          rec.calc = this.engine.cachedValue(sh.name, r, c);
        }
      }
    }
  }

  snapshotSheet(name: string): SheetSnapshot | undefined {
    const sh = this.getSheet(name);
    if (!sh) return undefined;
    const cells = new Map<number, CellRec>();
    for (const [k, v] of sh.cells) cells.set(k, cloneRec(v));
    return {
      name: sh.name, original: sh.original, metadata: sh.metadata, cells, rowCount: sh.rowCount, colCount: sh.colCount,
      colWidths: new Map(sh.colWidths), rowHeights: new Map(sh.rowHeights),
      hiddenRows: new Set(sh.hiddenRows), hiddenCols: new Set(sh.hiddenCols),
      frozenRows: sh.frozenRows, frozenCols: sh.frozenCols,
    };
  }

  restoreSheetSnapshot(snap: SheetSnapshot, index?: number): void {
    // remove current engine entries for this sheet name
    const cur = this.getSheet(snap.name);
    if (cur) for (const [k, rec] of cur.cells) {
      if (rec.f !== undefined) { const { r, c } = rcOf(k); this.engine.removeFormula(cur.name, r, c); }
    }
    const sh: SheetData = {
      name: snap.name, original: snap.original, metadata: snap.metadata, cells: new Map([...snap.cells].map(([k,v]) => [k, cloneRec(v)])), rowCount: snap.rowCount, colCount: snap.colCount,
      colWidths: new Map(snap.colWidths), rowHeights: new Map(snap.rowHeights),
      hiddenRows: new Set(snap.hiddenRows), hiddenCols: new Set(snap.hiddenCols),
      frozenRows: snap.frozenRows, frozenCols: snap.frozenCols,
    };
    const existing = this.sheets.findIndex((s) => s.name === snap.name);
    if (existing >= 0) this.sheets[existing] = sh;
    else this.sheets.splice(index ?? this.sheets.length, 0, sh);
    for (const [k, rec] of sh.cells) {
      if (rec.f !== undefined) { const { r, c } = rcOf(k); rec.calc = this.engine.setFormula(sh.name, r, c, rec.f); }
    }
    this.rebuildFormulas();
    this.dirty = true;
  }

  // ---------- worksheets ----------

  uniqueSheetName(base: string): string {
    base = base.slice(0,31);
    const names = new Set(this.sheetNames().map(n => n.toLowerCase()));
    if (!names.has(base.toLowerCase())) return base;
    let i = 2;
    let candidate: string;
    do { const suffix = ` (${i++})`; candidate = base.slice(0,31-suffix.length)+suffix; } while (names.has(candidate.toLowerCase()));
    return candidate;
  }

  addSheet(base = 'Sheet', index?: number): SheetData {
    const existing = this.sheets.filter((s) => s.name.startsWith(base)).length;
    const name = this.uniqueSheetName(existing ? `${base}${existing + 1}` : `${base}1`);
    const sh: SheetData = {
      name, cells: new Map(), rowCount: 100, colCount: 10,
      colWidths: new Map(), rowHeights: new Map(), hiddenRows: new Set(), hiddenCols: new Set(), frozenRows: 0, frozenCols: 0,
    };
    this.sheets.splice(index ?? this.sheets.length, 0, sh);
    this.dirty = true;
    return sh;
  }

  deleteSheet(name: string): { snap: SheetSnapshot; index: number } | undefined {
    const idx = this.sheets.findIndex((s) => s.name === name);
    if (idx < 0) return undefined;
    if (this.sheets.length <= 1) throw new Error('A workbook must contain at least one sheet.');
    const snap = this.snapshotSheet(name)!;
    for (const [k, rec] of this.sheets[idx].cells) {
      if (rec.f !== undefined) { const { r, c } = rcOf(k); this.engine.removeFormula(name, r, c); }
    }
    this.sheets.splice(idx, 1);
    this.engine.removeSheet(name);
    this.rebuildFormulas();
    this.dirty = true;
    return { snap, index: idx };
  }

  renameSheet(oldName: string, newName: string): boolean {
    newName = newName.trim();
    if (newName === oldName) return false;
    if (!newName || newName.startsWith("'") || newName.endsWith("'")) throw new Error('Invalid worksheet name.');
    if (this.sheets.some((s) => s !== this.getSheet(oldName) && s.name.toLowerCase() === newName.toLowerCase())) throw new Error('Worksheet name already exists.');
    if (/[\\/:?*[\]]/.test(newName) || newName.length > 31) throw new Error('Invalid worksheet name.');
    const sh = this.getSheet(oldName);
    if (!sh) return false;
    sh.name = newName;
    this.engine.renameSheet(oldName, newName);
    // rewrite cross-sheet refs textually (best effort)
    for (const s of this.sheets) {
      for (const [, rec] of s.cells) {
        if (rec.f !== undefined) {
          rec.f = renameFormulaSheet(rec.f, oldName, newName);
        }
      }
    }
    this.rebuildFormulas();
    this.dirty = true;
    return true;
  }

  duplicateSheet(name: string): SheetData | undefined {
    const src = this.getSheet(name);
    if (!src) return undefined;
    const snap = this.snapshotSheet(name)!;
    const copy = this.uniqueSheetName(`${name} (2)`);
    snap.name = copy;
    const idx = this.sheets.findIndex((s) => s.name === name) + 1;
    this.restoreSheetSnapshot(snap, idx);
    return this.getSheet(copy);
  }

  moveSheet(name: string, toIndex: number): void {
    const from = this.sheets.findIndex((s) => s.name === name);
    if (from < 0) return;
    toIndex = Math.max(0, Math.min(this.sheets.length - 1, toIndex));
    const [sh] = this.sheets.splice(from, 1);
    this.sheets.splice(toIndex, 0, sh);
    this.dirty = true;
  }

  // ---------- serialization (preserve original workbook) ----------

  toBytes(format = this.fileType, activeSheet = this.sheets[0]?.name): Uint8Array {
    if (!['xlsx','xls','xlsm','xlsb','csv'].includes(format)) throw new Error('Unsupported output format.');
    if (!this.sheets.length) throw new Error('A workbook needs at least one worksheet.');
    if (['xls','xlsb'].includes(format) && this.sheets.some(sh=>[...sh.cells.values()].some(c=>c.f!==undefined))) throw new Error('SheetJS cannot write formulas to XLS/XLSB. Choose XLSX or XLSM to preserve formulas.');
    if (format === 'csv' && this.sheets.length !== 1) throw new Error('CSV can only save one sheet. Save as XLSX to retain all sheets.');
    if (this.wb?.vbaraw && !['xlsm','xlsb','xls'].includes(format)) throw new Error('Choose XLSM, XLSB or XLS to preserve VBA macros.');
    // Serialize into new containers. A failed write must never mutate our baseline.
    const wb: XLSX.WorkBook = { ...this.wb, SheetNames: this.sheetNames(), Sheets: {} };
    wb.Workbook = { ...this.wb?.Workbook, Sheets: this.sheets.map(sh => ({ ...sh.metadata, name: sh.name })) };
    for (const sh of this.sheets) {
      const ws: XLSX.WorkSheet = {};
      for (const [k,v] of Object.entries(sh.original ?? {})) if (k.startsWith('!')) ws[k] = v;
      let maxR = 0, maxC = 0;
      for (const [k,rec] of sh.cells) {
        const {r,c} = rcOf(k);
        maxR = Math.max(maxR,r); maxC = Math.max(maxC,c);
        const cell: XLSX.CellObject = { ...(rec.original ?? {t:'z'}) };
        if (rec.edited || rec.f !== undefined || !rec.original) {
          delete cell.w; delete cell.h; delete cell.r; delete cell.f;
          const val = rec.f !== undefined ? rec.calc ?? rec.v : rec.v;
          if (rec.f !== undefined) cell.f = rec.f;
          if (isErr(val)) {
            // Error cells require Excel's numeric error codes, never error strings.
            const codes: Record<string,number> = {'#NULL!':0,'#DIV/0!':7,'#VALUE!':15,'#REF!':23,'#NAME?':29,'#NUM!':36,'#N/A':42,'#CYCLE!':23};
            if (!rec.edited && rec.original?.f && (val.err === '#NAME?' || val.err === '#VALUE!')) {
              cell.t = rec.original.t; cell.v = rec.original.v;
            } else { cell.t = 'e'; cell.v = codes[val.err] ?? 15; }
          } else if (typeof val === 'number') { cell.t='n'; cell.v=val; }
          else if (typeof val === 'boolean') { cell.t='b'; cell.v=val; }
          else if (typeof val === 'string') { cell.t='s'; cell.v=val; }
          else { cell.t='z'; delete cell.v; }
        }
        const style = styleToXlsx(rec.style);
        if (style) cell.s = style;
        if (rec.style?.numFmt) cell.z = rec.style.numFmt;
        ws[cellAddr(r,c)] = cell;
      }
      ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR,c:maxC}});
      const cols: XLSX.ColInfo[] = (sh.original?.['!cols'] ?? []).map(c => ({...c,hidden:false,width:undefined,wpx:undefined,wch:undefined}));
      for (const [i,w] of sh.colWidths) cols[i] = {...cols[i],width:undefined,wpx:undefined,wch:w,MDW:7};
      for (const i of sh.hiddenCols) cols[i] = {...cols[i],hidden:true};
      ws['!cols'] = cols;
      const rows: XLSX.RowInfo[] = (sh.original?.['!rows'] ?? []).map(r => ({...r,hidden:false,hpx:undefined,hpt:undefined}));
      for (const [i,h] of sh.rowHeights) rows[i] = {...rows[i],hpt:undefined,hpx:h};
      for (const i of sh.hiddenRows) rows[i] = {...rows[i],hidden:true};
      ws['!rows'] = rows;
      wb.Sheets[sh.name] = ws;
    }
    const out = XLSX.write(wb, {type:'array',bookType:format as XLSX.BookType,bookVBA:true,cellStyles:true,compression:true,sheet:activeSheet});
    return out instanceof Uint8Array ? out : new Uint8Array(out as ArrayBuffer);
  }

  /** Row data for AG Grid: materializes visible rows for one sheet (virtualized by grid). */
  gridRows(sheet: string, uptoRow?: number): { id: number; cells: Record<string, { v: RawVal; text: string; formula?: string; style?: CellStyle }> }[] {
    const sh = this.getSheet(sheet);
    if (!sh) return [];
    const total = Math.min(uptoRow ?? sh.rowCount, 200000);
    const byRow = new Map<number, { c: number; rec: CellRec }[]>();
    for (const [k, rec] of sh.cells) {
      const r = Math.floor(k / 16384), c = k % 16384;
      if (r >= total || sh.hiddenCols.has(c)) continue;
      let arr = byRow.get(r);
      if (!arr) { arr = []; byRow.set(r, arr); }
      arr.push({ c, rec });
    }
    const rows: { id: number; cells: Record<string, { v: RawVal; text: string; formula?: string; style?: CellStyle }> }[] = [];
    for (let r = 0; r < total; r++) {
      if (sh.hiddenRows.has(r)) continue;
      const cells: Record<string, { v: RawVal; text: string; formula?: string; style?: CellStyle }> = {};
      const arr = byRow.get(r);
      if (arr) {
        for (const { c, rec } of arr) {
          const d = this.displayOf(sheet, r, c);
          cells[String(c)] = { v: d.v, text: d.text, ...(d.formula !== undefined ? { formula: d.formula } : {}), ...(rec.style ? { style: rec.style } : {}) };
        }
      }
      rows.push({ id: r + 1, cells });
    }
    return rows;
  }
}

export function toBase64(bytes: Uint8Array): string {
  let s = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) {
    s += String.fromCharCode(...bytes.subarray(i, i + CH));
  }
  return btoa(s);
}

export function fromBase64(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
