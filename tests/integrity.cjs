const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const XLSX = require('xlsx');
const ts=require('typescript');
require.extensions['.ts']=(module,filename)=>{const source=fs.readFileSync(filename,'utf8');const js=ts.transpileModule(source,{compilerOptions:{target:ts.ScriptTarget.ES2022,module:ts.ModuleKind.CommonJS,esModuleInterop:true}}).outputText;module._compile(js,filename);};
const {WorkbookModel} = require('../src/workbook/model.ts');
const {editCellsCommand,insertDeleteCommand,sheetOpCommand,CommandManager,layoutCommand} = require('../src/workbook/commands.ts');
const {parseTSV}=require('../src/workbook/clipboard.ts');
let passed=0;
function test(name,fn){fn();passed++;console.log('PASS '+name);}
function fixture(){
  const wb=XLSX.utils.book_new();
  const sh=XLSX.utils.aoa_to_sheet([[1,2,3],[4,5,6],[7,8,9]]);
  sh.D1={t:'n',v:3,f:'A1+B1'};sh.E1={t:'e',v:7};sh.A1.c=[{a:'Reviewer',t:'Keep me'}];sh.B1.l={Target:'https://example.com'};
  sh.B2.z='0.00';sh['!cols']=[{hidden:true},{wch:18}];sh['!rows']=[{hpx:45},{hidden:true}];
  XLSX.utils.book_append_sheet(wb,sh,'Data');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([[10]]),'Other');
  wb.Props={Title:'Metadata survives'};
  wb.Workbook={Sheets:[{name:'Data',Hidden:0},{name:'Other',Hidden:1}]};
  sh['!ref']='A1:E3';
  return wb;
}
function load(wb=fixture(),format='xlsx'){const m=new WorkbookModel();m.loadFromBytes(XLSX.write(wb,{type:'buffer',bookType:format,cellStyles:true}),format);return m;}
function edit(m,s,r,c,v,f){const cmd=editCellsCommand(m,s,[{r,c,next:{v,f}}]);cmd.do();return cmd;}
test('round-trip types, comments, links, errors, hidden dimensions, sheet metadata',()=>{
  for(const format of ['xlsx','xlsm']){
    const m=load(fixture(),format);edit(m,'Data',1,1,22);
    const bytes=m.toBytes(format);const wb=XLSX.read(bytes,{type:'array',cellStyles:true,cellNF:true});
    assert.equal(wb.Sheets.Data.B2.v,22,format);assert.equal(wb.Sheets.Data.D1.f,'A1+B1');
    assert.equal(wb.Sheets.Data.E1.t,'e');assert.equal(wb.Sheets.Data.E1.v,7);
    assert.equal(wb.Sheets.Data.A1.c[0].t,'Keep me');assert.equal(wb.Sheets.Data.B1.l.Target,'https://example.com');
    assert.equal(wb.Sheets.Data['!cols'][0].hidden,true);assert.equal(wb.Workbook.Sheets[1].Hidden,1);
    assert.equal(wb.Sheets.Data.B2.z,'0.00');
    if(format==='xls')assert.equal(bytes[0],0xd0);else assert.equal(bytes[0],0x50);
  }
});
test('legacy exports reject formula loss and write correct binary formats',()=>{
  for(const format of ['xls','xlsb']){const m=load();assert.throws(()=>m.toBytes(format),/cannot write formulas/);edit(m,'Data',0,3,3);const bytes=m.toBytes(format);const wb=XLSX.read(bytes);assert.equal(wb.Sheets.Data.B2.v,5);assert.equal(bytes[0],format==='xls'?0xd0:0x50);}
});
test('serialize does not mutate original; rename retains sheet properties',()=>{
  const m=load();const before=JSON.stringify(m.wb);
  sheetOpCommand(m,{type:'rename',oldName:'Data',newName:'Renamed'}).do();
  const wb=XLSX.read(m.toBytes(),{type:'array',cellStyles:true});
  assert.equal(wb.Sheets.Renamed['!cols'][0].hidden,true);assert.equal(JSON.stringify(m.wb),before);
});
test('formula graph: batch paste, case-insensitive refs, long chains, cycles, sparse INDEX',()=>{
  const m=load();
  const changes=[{r:5,c:0,next:{v:null,f:'B6*2'}},{r:5,c:1,next:{v:11}}];
  editCellsCommand(m,'Data',changes).do();assert.equal(m.displayOf('Data',5,0).v,22);
  edit(m,'Other',0,1,null,'data!B6');edit(m,'Data',5,1,12);assert.equal(m.displayOf('Other',0,1).v,12);
  edit(m,'Data',9,0,null,'B10');edit(m,'Data',9,1,null,'A10');assert.equal(m.displayOf('Data',9,0).text,'#CYCLE!');
  for(let r=110;r>=20;r--)edit(m,'Data',r,2,r===110?9:null,r===110?undefined:`C${r+2}`);
  assert.equal(m.displayOf('Data',20,2).v,9);
  edit(m,'Other',5,0,90);edit(m,'Other',3,1,null,'INDEX(A4:A6,3)');assert.equal(m.displayOf('Other',3,1).v,90);
  edit(m,'Data',0,9,null,'UNSUPPORTED(A1)');assert.equal(m.getRec('Data',0,9).f,'UNSUPPORTED(A1)');
  const wb=XLSX.read(m.toBytes(),{type:'array'});assert.equal(wb.Sheets.Data.J1.f,'UNSUPPORTED(A1)');assert.equal(wb.Sheets.Data.J1.t,'e');assert.equal(typeof wb.Sheets.Data.J1.v,'number');
});
test('row/column delete undo restores deleted formulas, external dependents and dimensions',()=>{
  for(const axis of ['row','col']){
    const m=load();edit(m,'Other',0,1,null,'Data!$B$2');
    edit(m,'Other',0,2,null,'A1');edit(m,'Other',0,3,null,'"Data!B2"');
    const history=new CommandManager();history.exec(insertDeleteCommand(m,'Data',axis,'delete',1,1));
    assert.equal(m.getRec('Other',0,1).f,'#REF!');assert.equal(m.getRec('Other',0,2).f,'A1');assert.equal(m.getRec('Other',0,3).f,'"Data!B2"');
    history.undo();assert.equal(m.getRec('Other',0,1).f,'Data!$B$2');assert.equal(m.displayOf('Other',0,1).v,5);
    assert.equal(m.getSheet('Data').hiddenRows.has(1),true);assert.equal(m.getSheet('Data').hiddenCols.has(0),true);
    history.redo();history.undo();assert.equal(m.getRec('Data',1,1).v,5);
  }
});
test('insert shifts absolute refs and ranges; unrelated sheets and strings unchanged',()=>{
  const m=load();edit(m,'Other',0,1,null,'Data!$B$2');edit(m,'Other',0,2,null,'A2');edit(m,'Data',0,5,null,'SUM(A1:A3)');
  insertDeleteCommand(m,'Data','row','insert',1,1).do();
  assert.equal(m.getRec('Other',0,1).f,'Data!$B$3');assert.equal(m.getRec('Other',0,2).f,'A2');assert.equal(m.getRec('Data',0,5).f,'SUM(A1:A4)');
});
test('sheet delete undo, rename literals, duplicate snapshot isolation and final sheet guard',()=>{
  const m=load();edit(m,'Other',0,1,null,'Data!A1');edit(m,'Other',0,2,null,'"Data!A1"');
  const rename=sheetOpCommand(m,{type:'rename',oldName:'Data',newName:'Sales 2026'});rename.do();
  assert.equal(m.getRec('Other',0,1).f,"'Sales 2026'!A1");assert.equal(m.getRec('Other',0,2).f,'"Data!A1"');rename.undo();
  const del=sheetOpCommand(m,{type:'delete',name:'Data'});del.do();assert.equal(m.getRec('Other',0,1).f,'#REF!');del.undo();assert.equal(m.displayOf('Other',0,1).v,1);
  const snap=m.snapshotSheet('Data');m.restoreSheetSnapshot(snap);edit(m,'Data',0,0,88);assert.equal(snap.cells.get(0).v,1);
  m.deleteSheet('Other');assert.throws(()=>m.deleteSheet('Data'),/at least one/);
});
test('history save point, branching, dimension undo and bounds validation',()=>{
  const m=load();const h=new CommandManager();h.exec(editCellsCommand(m,'Data',[{r:0,c:0,next:{v:2}}]));h.markClean();h.exec(editCellsCommand(m,'Data',[{r:0,c:0,next:{v:3}}]));h.undo();assert.equal(h.isDirty,false);h.redo();assert.equal(h.isDirty,true);
  const sh=m.getSheet('Data');h.exec(layoutCommand(m,'Data',()=>sh.frozenRows=1));h.undo();assert.equal(sh.frozenRows,0);h.redo();assert.equal(sh.frozenRows,1);
  assert.throws(()=>editCellsCommand(m,'Data',[{r:0,c:16384,next:{v:5}}]),/bounds/);
  assert.throws(()=>insertDeleteCommand(m,'Data','row','delete',0,999),/bounds/);
});
test('absolute formulas and error propagation',()=>{
 const m=load();edit(m,'Data',4,4,null,'$A$1+B$1+$B2');assert.equal(m.displayOf('Data',4,4).v,8);edit(m,'Data',4,5,null,'E1+1');assert.equal(m.displayOf('Data',4,5).text,'#DIV/0!');
});
test('clipboard quotes, multiline fields and CSV output',()=>{
  assert.deepEqual(parseTSV('a"b\t"x\ny"\r\n2\t3'),[['a"b','x\ny'],['2','3']]);
  const m=new WorkbookModel();m.loadFromBytes(new TextEncoder().encode('name,value\r\nAlice,2'),'csv');edit(m,m.sheetNames()[0],1,1,9);
  const wb=XLSX.read(m.toBytes('csv'),{type:'array'});assert.equal(wb.Sheets[wb.SheetNames[0]].B2.v,9);
  assert.throws(()=>load().toBytes('csv'),/one sheet/);
});
async function storeTests(){
  const bytes=XLSX.write(fixture(),{type:'buffer',bookType:'xlsx'}).toString('base64');
  let resolveSave;let reads=0;let formatUsed;let fail=false;
  global.window={confirm:()=>true,viewer:{
    openPath:async p=>({ok:true,data:{name:'original.xlsx',path:p,sheets:['Data','Other'],size:10}}),
    readBytes:async()=>{reads++;return{ok:true,data:{path:'C:/original.xlsx',base64:bytes}};},
    recent:async()=>({ok:true,data:[]}),
    chooseSavePath:async()=>({ok:true,data:{path:'C:/copy.xlsm',format:'xlsm'}}),
    save:async(b,p,f)=>{formatUsed=f;await new Promise(r=>resolveSave=r);return fail?{ok:false,error:'disk full'}:{ok:true,data:{path:p}};}
  }};
  const {useViewer}=require('../src/store.ts');const st=()=>useViewer.getState();
  await st().open('C:/original.xlsx');const original=st().model;
  st().exec(editCellsCommand(original,'Data',[{r:0,c:0,next:{v:99}}]));
  const saving=st().saveAs();await new Promise(r=>setImmediate(r));assert.equal(st().busy,true);
  st().exec(editCellsCommand(original,'Data',[{r:0,c:0,next:{v:500}}]));assert.equal(original.getRec('Data',0,0).v,99);
  resolveSave();assert.equal(await saving,true);assert.equal(st().dirty,false);assert.equal(st().model,original);assert.equal(reads,1);assert.equal(formatUsed,'xlsm');
  st().undo();assert.equal(st().dirty,true);st().redo();assert.equal(st().dirty,false);
  st().exec(editCellsCommand(original,'Data',[{r:0,c:0,next:{v:100}}]));fail=true;
  const failed=st().save();await new Promise(r=>setImmediate(r));resolveSave();assert.equal(await failed,false);assert.equal(st().dirty,true);assert.equal(st().filePath,'C:/copy.xlsm');
  const all=[...original.getSheet('Data').cells.keys()].map(k=>({r:Math.floor(k/16384),c:k%16384,next:{v:null}}));
  st().exec(editCellsCommand(original,'Data',all));await st().load('Other');await st().load('Data');assert.equal(original.getRec('Data',0,0)?.v??null,null);
  const choose=window.viewer.chooseSavePath;
  delete window.viewer.chooseSavePath;
  assert.equal(await st().saveAs(),false);
  assert.match(st().error,/older session/);
  assert.equal(st().dirty,true);
  window.viewer.chooseSavePath=choose;
  console.log('PASS store: no reload, busy lock, save failure, correct format, save point, stale bridge and cleared sheet switching');
}
storeTests().then(()=>console.log(`${passed+1} integrity groups passed`)).catch(e=>{console.error(e);process.exitCode=1;});
