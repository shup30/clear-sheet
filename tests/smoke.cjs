const { _electron: electron } = require('playwright');
const XLSX = require('xlsx');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {Worker} = require('node:worker_threads');
const out = path.resolve('test-output');
fs.mkdirSync(out,{recursive:true});
const book=XLSX.utils.book_new();
const sheet=XLSX.utils.aoa_to_sheet([['Name','Amount','Total'],['Alice',12,24],['Bob',8,16],['Carol',30,60]]);
sheet.C2.f='B2*2';
XLSX.utils.book_append_sheet(book,sheet,'Sales');
XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Second sheet'],['Lazy loaded']]),'Notes');
for(const ext of ['xlsx','xls','xlsm','xlsb','csv']) XLSX.writeFile(book,path.join(out,'sample.'+ext));
fs.writeFileSync(path.join(out,'broken.xlsx'),'PK\x03\x04invalid compressed workbook');
const file=path.join(out,'sample.xlsx');
async function main(){
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
  env.CLEAR_SHEET_TEST_USER_DATA=path.join(out,'profile-smoke.cjs-'+process.pid);fs.mkdirSync(env.CLEAR_SHEET_TEST_USER_DATA,{recursive:true});
 const app=await electron.launch({args:[path.resolve('.'),file],env});
 try {
  const page=await app.firstWindow();
  const errors=[];page.on('pageerror',e=>errors.push(e.message));
  await page.getByRole('tab',{name:'Sales',exact:true}).waitFor();
  await page.getByText('Alice',{exact:true}).waitFor();
  assert.equal(await page.evaluate(()=>typeof window.require),'undefined');
  assert.equal(await page.evaluate(()=>typeof window.process),'undefined');
  await page.locator('.ag-row[row-index="1"] [col-id="2"]').click();
  assert.equal(await page.getByRole('textbox',{name:'Formula bar'}).inputValue(),'=B2*2');
  await page.getByRole('button',{name:'Copy cell',exact:true}).click();
  assert.equal(await app.evaluate(({clipboard})=>clipboard.readText()),'24');
   await page.getByRole('textbox',{name:'Search sheet'}).fill('Bob');
   await page.getByText('Bob',{exact:true}).waitFor();
   await page.getByText('Alice',{exact:true}).waitFor({state:'detached'});
  await page.getByRole('button',{name:'Reset filters'}).click();
  await page.getByText('Alice',{exact:true}).waitFor();
  await page.getByRole('tab',{name:'Notes',exact:true}).click();
  await page.getByText('Lazy loaded',{exact:true}).waitFor();
  await page.getByRole('tab',{name:'Sales',exact:true}).click();
  await page.getByText('Alice',{exact:true}).waitFor();
  // Numeric sort keeps the original spreadsheet row address.
  await page.locator('.ag-header-cell[col-id="1"] .ag-header-cell-text').click();
  await page.locator('.ag-row[row-index="0"] [col-id="0"]').filter({hasText:'Bob'}).waitFor();
  await page.locator('.ag-row[row-index="0"] [col-id="1"]').click();
  assert.equal(await page.locator('.address').innerText(),'B3');
  await page.screenshot({path:path.join(out,'viewer.png')});
  for(const ext of ['xlsx','xls','xlsm','xlsb','csv']) {
   const result=await page.evaluate(async p=>{const b=await window.viewer.openPath(p);if(!b.ok)return b;return window.viewer.sheet(b.data.sheets[0]);},path.join(out,'sample.'+ext));
   assert.equal(result.ok,true,ext+': '+result.error);
   assert.equal(result.data.rows[1].cells['0'].text,'Alice');
   console.log('PASS format '+ext);
  }
  const bad=await page.evaluate(p=>window.viewer.openPath(p),path.join(out,'broken.xlsx'));
  assert.equal(bad.ok,false);
  const unsupported=await page.evaluate(p=>window.viewer.openPath(p),path.join(out,'sample.exe'));
  assert.equal(unsupported.ok,false);
  const recent=await page.evaluate(()=>window.viewer.recent());
  assert.equal(recent.ok,true);assert.ok(recent.data.includes(file));
  assert.deepEqual(errors,[]);
  console.log('PASS UI, startup file, formula, copy, search, numeric sort, tabs, recent files, invalid input, renderer isolation');
 } finally {await app.close();}
 // Check the preview cap without passing a million cells to the renderer.
 const huge=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([['Start']]);ws['!ref']='A1:A250000';
 XLSX.utils.book_append_sheet(huge,ws,'Large');XLSX.writeFile(huge,path.join(out,'large.xlsx'));
 const worker=new Worker(path.resolve('dist-electron/worker.js'));
 const request=(action,args)=>new Promise((resolve,reject)=>{worker.once('message',m=>m.error?reject(new Error(m.error)):resolve(m.data));worker.once('error',reject);worker.postMessage({id:1,action,...args});});
 try{await request('open',{path:path.join(out,'large.xlsx')});const data=await request('sheet',{name:'Large'});assert.equal(data.truncated,true);assert.equal(data.rows.length,200000);console.log('PASS large-sheet preview limit');}finally{await worker.terminate();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});
