import { requireSaveBridge } from './desktop';
import { create } from 'zustand';
import type { Book, Sheet, ViewerAPI, Result } from '../electron/types';
import { WorkbookModel, fromBase64, toBase64 } from './workbook/model';
import { CommandManager, type Command } from './workbook/commands';
import { colToLetter } from './workbook/cellRef';

declare global { interface Window { viewer: ViewerAPI } }

function unwrap<T>(r: Result<T>): T { if (!r.ok) throw new Error(r.error); return r.data; }

export interface SelRange { r1: number; c1: number; r2: number; c2: number }

export interface Selection {
  active: { r: number; c: number };
  ranges: SelRange[];
}

export interface RowData { id: number; cells: Record<string, { v: string | number | boolean | null; text: string; formula?: string; style?: import('../electron/types').CellStyle }> }

interface State {
  book: Book | null;
  sheetName: string | null;
  /** legacy preview (kept for compat) */
  sheet: Sheet | null;
  busy: boolean;
  error: string;
  recents: string[];
  zoom: number;
  search: string;
  selected: { address: string; text: string; formula?: string } | null;
  selection: Selection | null;
  model: WorkbookModel | null;
  filePath: string | null;
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  /** bump to rebuild grid rows */
  rev: number;
  layoutRev: number;
  rows: RowData[];
  findOpen: boolean;
  open: (path?: string, file?: File) => Promise<void>;
  load: (name: string) => Promise<void>;
  refresh: () => Promise<void>;
  set: (s: Partial<State>) => void;
  // editor
  ensureEditable: () => Promise<void>;
  exec: (cmd: Command) => void;
  undo: () => void;
  redo: () => void;
  setSelection: (sel: Selection | null) => void;
  markClean: (filePath?: string) => void;
  save: () => Promise<boolean>;
  saveAs: () => Promise<boolean>;
  confirmDiscard: () => boolean;
}

function buildRows(model: WorkbookModel, sheet: string): RowData[] {
  return model.gridRows(sheet) as RowData[];
}

function addr(r: number, c: number): string { return `${colToLetter(c)}${r + 1}`; }

export const useViewer = create<State>((set, get) => {
  const cmd = new CommandManager();
  let indexedRows: RowData[] | undefined;
  let rowIndex=new Map<number,number>();
  function refreshView(cmdObj: Command) {
      const st = get();
      if(cmdObj.viewChange==='dimensions'){set({rev:st.rev+1,layoutRev:st.layoutRev+1,dirty:cmd.isDirty,canUndo:cmd.canUndo,canRedo:cmd.canRedo});return;}
      if (st.model && st.sheetName) {
        const t = cmdObj.touched;
        if (t && t.length && t.length <= 5000 && cmdObj.sheets.length === 1 && cmdObj.sheets[0] === st.sheetName) {
          // Incremental patch: update only touched rows in place.
          const rows = [...st.rows];
          if(indexedRows!==st.rows){ rowIndex=new Map(st.rows.map((row,i)=>[row.id,i])); }
          const idx=rowIndex;
          indexedRows=rows;
          const sh = st.model.getSheet(st.sheetName);
          let expanded = false;
          for (const cell of t) {
            if (cell.sheet !== st.sheetName || !sh) continue;
            if (sh.hiddenRows.has(cell.r) || sh.hiddenCols.has(cell.c)) continue;
            const ri = idx.get(cell.r + 1);
            const d = st.model.displayOf(st.sheetName, cell.r, cell.c);
            const rec = st.model.getRec(st.sheetName, cell.r, cell.c);
            const entry = { v: d.v, text: d.text, ...(d.formula !== undefined ? { formula: d.formula } : {}), ...(rec?.style ? { style: rec.style } : {}) };
            if (ri === undefined) { if(cell.r>=200000)continue;expanded = true; break; }
            const copy = { ...rows[ri], cells: { ...rows[ri].cells } };
            if (d.v === null && d.formula === undefined && !rec?.style) delete copy.cells[String(cell.c)];
            else copy.cells[String(cell.c)] = entry;
            rows[ri] = copy;
          }
          if (expanded) {
            set({ rows: buildRows(st.model, st.sheetName), rev: st.rev + 1, dirty: cmd.isDirty, canUndo: cmd.canUndo, canRedo: cmd.canRedo });
          } else {
            set({ rows, rev: st.rev + 1, dirty: cmd.isDirty, canUndo: cmd.canUndo, canRedo: cmd.canRedo });
          }
          // refresh formula-bar selection text
          const sel = get().selection;
          if (sel && st.model && st.sheetName) {
            const dd = st.model.displayOf(st.sheetName, sel.active.r, sel.active.c);
            set({ selected: { address: addr(sel.active.r, sel.active.c), text: dd.text, formula: dd.formula } });
          }
        } else {
          const names = st.model.sheetNames();
          const cur = st.sheetName && names.includes(st.sheetName) ? st.sheetName : names[0];
          set({
            book: st.book ? { ...st.book, sheets: names } : st.book,
            sheetName: cur ?? null,
            rows: cur ? buildRows(st.model, cur) : [],
            rev: st.rev + 1, layoutRev: st.layoutRev+1, dirty: cmd.isDirty,
            canUndo: cmd.canUndo, canRedo: cmd.canRedo,
          });
        }
      } else {
        set({ dirty: cmd.isDirty, canUndo: cmd.canUndo, canRedo: cmd.canRedo });
      }
      get().setSelection(get().selection);
  }
  return {
    book: null, sheet: null, sheetName: null, busy: false, error: '', recents: [],
    zoom: 100, search: '', selected: null, selection: null,
    model: null, filePath: null, dirty: false, canUndo: false, canRedo: false,
    rev: 0, layoutRev: 0, rows: [], findOpen: false,
    set: (s) => set(s),

    refresh: async () => {
      try { set({ recents: unwrap(await window.viewer.recent()) }); }
      catch (e) { set({ error: String(e) }); }
    },

    open: async (path, file) => {
      if (get().busy || (get().dirty && !get().confirmDiscard())) return;
      set({busy:true,error:''});
      try {
        const book=unwrap(await (file ? window.viewer.drop(file) : path ? window.viewer.openPath(path) : window.viewer.open()));
        if (!book) return;
        const data=unwrap(await window.viewer.readBytes(book.path));
        const model=new WorkbookModel();
        model.loadFromBytes(fromBase64(data.base64),data.path.split('.').pop() ?? 'xlsx');
        model.sourcePath=data.path;
        const first=model.sheetNames()[0];
        const rows=buildRows(model,first);
        // Commit only after a complete load. A preview can never become saveable.
        cmd.clear();
        set({book:{...book,sheets:model.sheetNames()},model,filePath:data.path,sheet:null,sheetName:first,rows,selection:null,selected:null,search:'',dirty:false,canUndo:false,canRedo:false,rev:get().rev+1});
        await get().refresh();
      } catch(e) {set({error:e instanceof Error?e.message:String(e)});}
      finally {set({busy:false});}
    },

    load: async name => {
      const st=get();
      if(st.busy || !st.model?.getSheet(name))return;
      // Never re-import worker previews over cleared or newly created sheets.
      set({sheetName:name,rows:buildRows(st.model,name),sheet:null,selected:null,selection:null,search:'',rev:st.rev+1});
    },

    ensureEditable: async () => {
      if (!get().model?.wb) throw new Error('Workbook is not fully loaded. Reopen it before editing or saving.');
    },

    exec: (cmdObj) => {
      if (get().busy || !get().model?.wb) return;
      try { cmd.exec(cmdObj); } catch(e) {set({error:e instanceof Error?e.message:String(e)});return;}
      refreshView(cmdObj);
    },

    undo: () => {
      if(get().busy)return;
      try {const changed=cmd.undo();if(changed)refreshView(changed);}catch(e){set({error:String(e)});}
    },
    redo: () => {
      if(get().busy)return;
      try {const changed=cmd.redo();if(changed)refreshView(changed);}catch(e){set({error:String(e)});}
    },

    setSelection: (sel) => {
      if (!sel) { set({ selection: null, selected: null }); return; }
      const st = get();
      const { r, c } = sel.active;
      let text = '', formula: string | undefined;
      if (st.model && st.sheetName) {
        const d = st.model.displayOf(st.sheetName, r, c);
        text = d.text; formula = d.formula;
      }
      set({ selection: sel, selected: { address: addr(r, c), text, formula } });
    },

    markClean: (filePath) => {
      if (filePath) set({ filePath });
      cmd.markClean();
      if (get().model) get().model!.dirty = false;
      set({ dirty: false });
    },

    confirmDiscard: () => {
      // Synchronous confirm so menu/open flows can abort.
      return window.confirm('You have unsaved changes. Discard them?');
    },

    save: async () => {
      const st=get();
      if(st.busy)return false;
      if(!st.model?.wb || !st.sheetName) {set({error:'Nothing to save.'});return false;}
      if(!st.filePath || st.model.requiresCopy) return get().saveAs();
      set({busy:true,error:''});
      try {
        const format=st.filePath.split('.').pop()!.toLowerCase();
        const bytes=st.model.toBytes(format,st.sheetName);
        const result=unwrap(await window.viewer.save(toBase64(bytes),st.filePath,format));
        get().markClean(result.path);
        await get().refresh();
        return true;
      } catch(e){set({error:e instanceof Error?e.message:String(e)});return false;}
      finally{set({busy:false});}
    },

    saveAs: async () => {
      const st=get();
      if(st.busy)return false;
      if(!st.model?.wb) {set({error:'Nothing to save.'});return false;}
      set({busy:true,error:''});
      try {
        const suggested=st.model.requiresCopy ? (st.filePath ?? 'workbook.xlsx').replace(/(\.[^.]+)$/, ' - edited$1') : st.filePath ?? 'workbook.xlsx';
        requireSaveBridge(window.viewer);
        const target=unwrap(await window.viewer.chooseSavePath(suggested));
        if(!target)return false;
        if(st.model.requiresCopy && target.path.toLowerCase()===st.model.sourcePath?.toLowerCase()) throw new Error('Choose a different file name. Excel features not supported by SheetJS must be preserved in the original.');
        const bytes=st.model.toBytes(target.format,st.sheetName ?? undefined);
        const result=unwrap(await window.viewer.save(toBase64(bytes),target.path,target.format));
        st.model.fileType=target.format;
        st.model.requiresCopy=false;
        get().markClean(result.path);
        set({book:st.book?{...st.book,name:result.path.split(/[/\\]/).pop()!,path:result.path}:null});
        await get().refresh();
        return true;
      } catch(e){set({error:e instanceof Error?e.message:String(e)});return false;}
      finally{set({busy:false});}
    },
  };
});
