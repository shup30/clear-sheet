import { parentPort } from 'node:worker_threads';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import type { Row, Sheet } from './types';
let buffer:Buffer;
let names:string[]=[];
let cache:Sheet|undefined;
parentPort!.on('message', ({id,action,path,name})=>{
 try {
  if(action==='open') {
   buffer=readFileSync(path);
   const book=XLSX.read(buffer,{type:'buffer',bookSheets:true,bookProps:true});
   names=book.SheetNames;
   if(!names.length) throw new Error('This file contains no worksheets.');
   cache=undefined;
   parentPort!.postMessage({id,data:names}); return;
  }
  if(!names.includes(name)) throw new Error('Unknown worksheet.');
  if(cache?.name===name) {parentPort!.postMessage({id,data:cache}); return;}
   const book=XLSX.read(buffer,{type:'buffer',sheets:name,cellFormula:true,cellText:true,sheetRows:200001});
   const ws=book.Sheets[name];
   const range=XLSX.utils.decode_range(ws['!ref']||'A1');
   const full=XLSX.utils.decode_range(ws['!fullref']||ws['!ref']||'A1');
   // `!ref` covers only rows parsed under the sheetRows cap; `!fullref` (when
   // present) carries the true sheet dimensions. Size the preview from the
   // full dimensions and leave unparsed cells blank.
   const columns=Math.min(Math.max(range.e.c,full.e.c)+1,1024);
   const count=Math.min(full.e.r+1,200000,Math.floor(1000000/columns));
  const rows:Row[]=[];
  for(let r=0;r<count;r++) {
   const cells:Row['cells']={};
   for(let c=0;c<columns;c++) {
    const cell=ws[XLSX.utils.encode_cell({r,c})];
    if(cell) cells[String(c)]={v:cell.v??null,text:cell.w??String(cell.v??''),...(cell.f?{formula:cell.f}:{})};
   }
   rows.push({id:r+1,cells});
  }
  cache={name,rows,columns,totalRows:full.e.r+1,truncated:count<full.e.r+1||columns<full.e.c+1};
  parentPort!.postMessage({id,data:cache});
 } catch(e) {parentPort!.postMessage({id,error:e instanceof Error?e.message:'Unable to read spreadsheet.'});}
});
