import {create} from 'zustand';
import type {Book,Sheet,ViewerAPI,Result} from '../electron/types';
declare global {interface Window {viewer:ViewerAPI}}
function unwrap<T>(r:Result<T>):T {if(!r.ok)throw new Error(r.error);return r.data;}
interface State {book:Book|null;sheet:Sheet|null;busy:boolean;error:string;recents:string[];zoom:number;search:string;selected:{address:string;text:string;formula?:string}|null;open:(path?:string,file?:File)=>Promise<void>;load:(name:string)=>Promise<void>;refresh:()=>Promise<void>;set:(s:Partial<State>)=>void}
export const useViewer=create<State>((set,get)=>({
 book:null,sheet:null,busy:false,error:'',recents:[],zoom:100,search:'',selected:null,
 set:s=>set(s),
 refresh:async()=>{try{set({recents:unwrap(await window.viewer.recent())});}catch(e){set({error:String(e)});}},
 open:async(path,file)=>{
  if(get().busy)return;
  set({busy:true,error:''});
  try {
   const book=unwrap(await(file?window.viewer.drop(file):path?window.viewer.openPath(path):window.viewer.open()));
   if(book){set({book,sheet:null,selected:null,search:''});set({sheet:unwrap(await window.viewer.sheet(book.sheets[0]))});await get().refresh();}
  }catch(e){set({error:e instanceof Error?e.message:String(e)});}finally{set({busy:false});}
 },
 load:async name=>{
  if(get().busy)return;
  set({busy:true,error:''});
  try{set({sheet:unwrap(await window.viewer.sheet(name)),selected:null,search:''});}catch(e){set({error:e instanceof Error?e.message:String(e)});}finally{set({busy:false});}
 }
}));
