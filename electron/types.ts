export type Result<T> = {ok:true;data:T}|{ok:false;error:string};
export interface Book {name:string; sheets:string[]; size:number}
export interface Cell {v:string|number|boolean|null; text:string; formula?:string}
export interface Row {id:number; cells:Record<string,Cell>}
export interface Sheet {name:string; rows:Row[]; columns:number; totalRows:number; truncated:boolean}
export interface ViewerAPI {
 open:()=>Promise<Result<Book|null>>;
 openPath:(path:string)=>Promise<Result<Book>>;
 drop:(file:File)=>Promise<Result<Book>>;
 sheet:(name:string)=>Promise<Result<Sheet>>;
 recent:()=>Promise<Result<string[]>>;
 clearRecent:()=>Promise<Result<void>>;
 copy:(text:string)=>Promise<Result<void>>;
 onOpen:(callback:()=>void)=>()=>void;
}
