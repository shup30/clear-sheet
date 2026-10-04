import { WorkbookModel, keyOf, rcOf, type CellRec } from './model';
import type { CellStyle } from '../../electron/types';
import { parseInput } from './format';
import { renameFormulaSheet } from './formulas';

export interface Command {
  label: string;
  /** grid sheets affected (for view refresh) */
  sheets: string[];
  /** specific cells touched (enables incremental grid patching) */
  viewChange?: 'dimensions';
  touched?: { sheet: string; r: number; c: number }[];
  do(): void;
  undo(): void;
}

/** Command manager: stores deltas only (never deep-clones the workbook). */
export class CommandManager {
  undoStack: Command[] = [];
  redoStack: Command[] = [];

  private cleanHead: Command | undefined;
  private cleanReachable = true;
  get isDirty(): boolean { return !this.cleanReachable || this.undoStack.at(-1) !== this.cleanHead; }
  markClean(): void { this.cleanHead = this.undoStack.at(-1); this.cleanReachable = true; }
  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }

  exec(cmd: Command): void {
    cmd.do();
    this.undoStack.push(cmd);
    if (this.undoStack.length > 200) { const dropped=this.undoStack.shift(); if(!this.cleanHead || dropped===this.cleanHead)this.cleanReachable=false; }
    if (this.cleanHead && this.redoStack.includes(this.cleanHead)) this.cleanReachable = false;
    this.redoStack = [];
  }

  undo(): Command | null {
    const c = this.undoStack.at(-1);
    if (!c) return null;
    c.undo();
    this.undoStack.pop();
    this.redoStack.push(c);
    return c;
  }

  redo(): Command | null {
    const c = this.redoStack.at(-1);
    if (!c) return null;
    c.do();
    this.redoStack.pop();
    this.undoStack.push(c);
    return c;
  }

  clear(): void { this.undoStack = []; this.redoStack = []; this.cleanHead = undefined; this.cleanReachable = true; }
}

export interface CellChange { r: number; c: number; prev?: CellRec; next: { v: CellRec['v']; f?: string; numFmt?: string } }

export function editCellsCommand(model: WorkbookModel, sheet: string, changes: CellChange[], label = 'Edit cells'): Command {
  for (const ch of changes) if (!Number.isInteger(ch.r) || !Number.isInteger(ch.c) || ch.r < 0 || ch.c < 0 || ch.r >= 1048576 || ch.c >= 16384) throw new Error('Paste/edit exceeds Excel bounds.');
  const before: { r: number; c: number; prev?: CellRec }[] = changes.map((ch) => {
    const cur = model.getRec(sheet, ch.r, ch.c);
    if(cur?.original?.F)throw new Error('Editing an array-formula range is not supported.');
    return { r: ch.r, c: ch.c, prev: cur ? { ...cur, style: cur.style ? {...cur.style} : undefined } : undefined };
  });
  return {
    label, sheets: [sheet],
    do() {
      for (const ch of changes) { model.setCell(sheet, ch.r, ch.c, ch.next.v, ch.next.f);if(ch.next.numFmt)model.setStyle(sheet,ch.r,ch.c,{numFmt:ch.next.numFmt}); }
      const aff = model.afterCellsChanged(changes.map((ch) => ({ sheet, r: ch.r, c: ch.c })));
      this.touched = [...changes.map((ch) => ({ sheet, r: ch.r, c: ch.c })), ...aff.map((a) => ({ sheet: a.sheet, r: a.r, c: a.c }))];
    },
    undo() {
      for (let i = 0; i < changes.length; i++) model.restoreCell(sheet, changes[i].r, changes[i].c, before[i].prev);
      const affected=model.afterCellsChanged(changes.map((ch) => ({ sheet, r: ch.r, c: ch.c })));
      this.touched=[...changes.map(ch=>({sheet,r:ch.r,c:ch.c})),...affected];
    },
  };
}

export function inputToChange(text: string, asText=false): { v: CellRec['v']; f?: string; numFmt?: string } {
  return asText ? {v:text || null} : parseInput(text);
}

export function formatCommand(
  model: WorkbookModel, sheet: string,
  targets: { r: number; c: number }[],
  patch: Partial<CellStyle> | null,
  numFmt?: string,
  label = 'Format',
): Command {
  const full: Partial<CellStyle> = { ...(patch ?? {}), ...(numFmt !== undefined ? { numFmt } : {}) };
  const before = targets.map((t) => model.getRec(sheet, t.r, t.c)?.style);
  const clearing = patch === null && numFmt === undefined;
  return {
    label, sheets: [sheet],
    touched: targets.map((t) => ({ sheet, r: t.r, c: t.c })),
    do() {
      for (const t of targets) {
        if (clearing) model.setStyle(sheet, t.r, t.c, null);
        else model.setStyle(sheet, t.r, t.c, full);
      }
    },
    undo() {
      for (let i = 0; i < targets.length; i++) {
        const t = targets[i];
        const prev = before[i];
        model.setStyle(sheet, t.r, t.c, null);
        if (prev) model.setStyle(sheet, t.r, t.c, prev);
      }
    },
  };
}

export function insertDeleteCommand(
  model: WorkbookModel, sheet: string,
  axis: 'row' | 'col', mode: 'insert' | 'delete', at: number, count: number,
): Command {
  const sh = model.getSheet(sheet);
  if (!sh) throw new Error('Unknown worksheet.');
  const size = axis === 'row' ? sh.rowCount : sh.colCount;
  const max = axis === 'row' ? 1048576 : 16384;
  if (!Number.isInteger(at) || !Number.isInteger(count) || at < 0 || count < 1 || at > size || (mode === 'delete' && at+count > size) || (mode === 'insert' && size+count > max)) throw new Error('Row/column operation exceeds worksheet bounds.');
  if (model.wb?.Workbook?.Names?.length || sh.original?.['!merges']?.length || sh.original?.['!autofilter'] || [...sh.cells.values()].some(c => c.original?.F)) throw new Error('Structural edits on named ranges, merged cells, filters or array formulas are not supported.');
  const formulas = formulaSnapshot(model);
  const dimensions = {original:sh.original,rowCount:sh.rowCount,colCount:sh.colCount,colWidths:new Map(sh.colWidths),rowHeights:new Map(sh.rowHeights),hiddenRows:new Set(sh.hiddenRows),hiddenCols:new Set(sh.hiddenCols),frozenRows:sh.frozenRows,frozenCols:sh.frozenCols};
  let removed: Map<number,CellRec> | undefined;
  return {
    label:`${mode} ${axis}`,sheets:model.sheetNames(),
    do() {
      if(axis==='row') { if(mode==='insert')model.insertRows(sheet,at,count);else removed=model.deleteRows(sheet,at,count); }
      else { if(mode==='insert')model.insertCols(sheet,at,count);else removed=model.deleteCols(sheet,at,count); }
    },
    undo() {
      if(axis==='row') { if(mode==='insert')model.deleteRows(sheet,at,count);else model.insertRows(sheet,at,count); }
      else { if(mode==='insert')model.deleteCols(sheet,at,count);else model.insertCols(sheet,at,count); }
      if (mode==='delete' && removed) for(const [k,rec] of removed) { const {r,c}=rcOf(k);model.restoreCell(sheet,r,c,rec); }
      Object.assign(sh,dimensions,{colWidths:new Map(dimensions.colWidths),rowHeights:new Map(dimensions.rowHeights),hiddenRows:new Set(dimensions.hiddenRows),hiddenCols:new Set(dimensions.hiddenCols)});
      restoreFormulas(model,formulas);
    }
  };
}

function formulaSnapshot(model: WorkbookModel) {
  return model.sheets.flatMap(sh => [...sh.cells].filter(([,c])=>c.f !== undefined).map(([k,c])=>({sheet:sh.name,k,f:c.f!})));
}
function restoreFormulas(model: WorkbookModel, snapshot: ReturnType<typeof formulaSnapshot>) {
  for(const item of snapshot) { const rec=model.getSheet(item.sheet)?.cells.get(item.k); if(rec)rec.f=item.f; }
  model.rebuildFormulas();
}

/** Small metadata snapshots, never cell/workbook clones. */
export function layoutCommand(model: WorkbookModel, sheet: string, mutate: () => void): Command {
  const sh=model.getSheet(sheet)!;
  const before={colWidths:new Map(sh.colWidths),rowHeights:new Map(sh.rowHeights),hiddenRows:new Set(sh.hiddenRows),hiddenCols:new Set(sh.hiddenCols),frozenRows:sh.frozenRows,frozenCols:sh.frozenCols};
  return {label:'Sheet layout',sheets:[sheet],do(){mutate();model.dirty=true;},undo(){Object.assign(sh,before,{colWidths:new Map(before.colWidths),rowHeights:new Map(before.rowHeights),hiddenRows:new Set(before.hiddenRows),hiddenCols:new Set(before.hiddenCols)});model.dirty=true;}};
}

export function sheetOpCommand(
  model: WorkbookModel,
  op: { type: 'add'; name?: string } | { type: 'delete'; name: string } | { type: 'duplicate'; name: string } | { type: 'rename'; oldName: string; newName: string } | { type: 'move'; name: string; from: number; to: number },
): Command {
  if ((op.type==='rename' || op.type==='delete') && model.wb?.Workbook?.Names?.length) throw new Error('Rename/delete with defined names is not supported.');
  const formulas = formulaSnapshot(model);
  let created: string | undefined;
  let deleted: { snap: ReturnType<WorkbookModel['snapshotSheet']>; index: number } | undefined;
  const sheetsOf = () => model.sheetNames();
  return {
    label: `Sheet ${op.type}`,
    sheets: sheetsOf(),
    do() {
      if (op.type === 'add') {
        created = model.addSheet(op.name ?? 'Sheet').name;
        this.sheets = sheetsOf();
      } else if (op.type === 'delete') {
        const r = model.deleteSheet(op.name);
        if (r) deleted = { snap: r.snap, index: r.index };
        for(const sh of model.sheets) for(const rec of sh.cells.values()) if(rec.f!==undefined)rec.f=renameFormulaSheet(rec.f,op.name,null);
        model.rebuildFormulas();
        this.sheets = sheetsOf();
      } else if (op.type === 'duplicate') {
        const c = model.duplicateSheet(op.name);
        created = c?.name;
        this.sheets = sheetsOf();
      } else if (op.type === 'rename') {
        model.renameSheet(op.oldName, op.newName);
        this.sheets = sheetsOf();
      } else if (op.type === 'move') {
        model.moveSheet(op.name, op.to);
        this.sheets = sheetsOf();
      }
    },
    undo() {
      if (op.type === 'add' && created) model.deleteSheet(created);
      else if (op.type === 'delete' && deleted?.snap) model.restoreSheetSnapshot(deleted.snap, deleted.index);
      else if (op.type === 'duplicate' && created) model.deleteSheet(created);
      else if (op.type === 'rename') model.renameSheet(op.newName, op.oldName);
      else if (op.type === 'move') model.moveSheet(op.name, op.from);
      restoreFormulas(model,formulas);
      this.sheets = sheetsOf();
    },
  };
}

export function keyOfCell(r: number, c: number): number { return keyOf(r, c); }

/** One user action spanning multiple sheets remains one undo step. */
export function compositeCommand(commands: Command[], label: string): Command {
  return {label,sheets:[...new Set(commands.flatMap(c=>c.sheets))],do(){
    const done:Command[]=[];
    try{for(const command of commands){command.do();done.push(command);}}
    catch(e){for(const command of done.reverse())command.undo();throw e;}
    this.touched=commands.flatMap(c=>c.touched??[]);
  },undo(){for(const command of [...commands].reverse())command.undo();this.touched=commands.flatMap(c=>c.touched??[]);}};
}
