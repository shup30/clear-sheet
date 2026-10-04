export type Result<T> = {ok:true;data:T}|{ok:false;error:string};
export interface Book {name:string; sheets:string[]; size:number; path?:string}
export interface CellStyle {
  bold?:boolean; italic?:boolean; underline?:boolean;
  fontName?:string; fontSize?:number; color?:string; bg?:string;
  hAlign?:'left'|'center'|'right'; vAlign?:'top'|'middle'|'bottom';
  wrap?:boolean; numFmt?:string;
  border?:{top?:boolean;bottom?:boolean;left?:boolean;right?:boolean};
}
export interface Cell {v:string|number|boolean|null; text:string; formula?:string; style?:CellStyle}
export interface Row {id:number; cells:Record<string,Cell>}
export interface Sheet {name:string; rows:Row[]; columns:number; totalRows:number; truncated:boolean}
export interface ViewerAPI {
  readonly bridgeVersion: 2;
  open:()=>Promise<Result<Book|null>>;
  openPath:(path:string)=>Promise<Result<Book>>;
  drop:(file:File)=>Promise<Result<Book>>;
  sheet:(name:string)=>Promise<Result<Sheet>>;
  recent:()=>Promise<Result<string[]>>;
  clearRecent:()=>Promise<Result<void>>;
  copy:(text:string)=>Promise<Result<void>>;
  paste:()=>Promise<Result<string>>;
  onOpen:(callback:()=>void)=>()=>void;
  /** Current file path known to main (after open/saveAs). */
  getPath:()=>Promise<Result<string|null>>;
  /** Raw file bytes (base64) for the current workbook, for the renderer-side editor model. */
  readBytes:(path?:string)=>Promise<Result<{base64:string; path:string}>>;
  /** Atomically overwrite the current file (or given path) with base64 workbook bytes. */
  save:(base64:string, targetPath:string, format:string)=>Promise<Result<{path:string}>>;
  /** Native dialog authorizes a destination; serialize only after its format is known. */
  chooseSavePath:(suggestedName:string)=>Promise<Result<{path:string;format:string}|null>>;
}
