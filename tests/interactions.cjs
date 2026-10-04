const {_electron:electron}=require('playwright');
const XLSX=require('xlsx');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');

async function main(){
 const out=path.resolve('test-output');fs.mkdirSync(out,{recursive:true});
 const file=path.join(out,'interactions.xlsx');const saved=path.join(out,'interactions-saved.xlsx');
 const book=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Name','Value'],['Alice',10],['Bob',20],['Carol',30],['Dave',40]]),'Data');
 XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Other sheet']]),'Other');
 XLSX.writeFile(book,file);
 const profile=path.join(out,'interaction-profile-'+process.pid);fs.mkdirSync(profile,{recursive:true});
 const env={...process.env,CLEAR_SHEET_TEST_USER_DATA:profile};delete env.ELECTRON_RUN_AS_NODE;
 const packaged=process.argv.includes('--packaged');
 const app=await electron.launch({...(packaged?{executablePath:path.resolve('release/win-unpacked/Clear Sheet.exe')}:{}),args:[...(packaged?['--user-data-dir='+profile]:[path.resolve('.')]),file],env,timeout:60000});
 const page=await app.firstWindow();page.setDefaultTimeout(15000);
 page.on('dialog',dialog=>dialog.dismiss().catch(()=>{}));
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 const width=()=>page.locator('.ag-header-cell[col-id="1"]').evaluate(el=>el.getBoundingClientRect().width);
 const rowHeight=()=>page.locator('.ag-center-cols-container .ag-row[row-id="2"]').evaluate(el=>el.getBoundingClientRect().height);
 const waitWidth=async expected=>page.waitForFunction(n=>Math.abs(document.querySelector('.ag-header-cell[col-id="1"]').getBoundingClientRect().width-n)<2,expected);
 const waitHeight=async expected=>page.waitForFunction(n=>Math.abs(document.querySelector('.ag-center-cols-container .ag-row[row-id="2"]').getBoundingClientRect().height-n)<2,expected);
 const cell=(id,col)=>page.locator(`.ag-center-cols-container .ag-row[row-id="${id}"] [col-id="${col}"]`);
 const undo=()=>page.getByRole('button',{name:'Undo',exact:true}).click();
 const redo=()=>page.getByRole('button',{name:'Redo',exact:true}).click();
 try {
  await page.getByText('Alice',{exact:true}).waitFor();
  assert.equal(await page.evaluate(()=>typeof window.viewer.chooseSavePath),'function');
  assert.equal(await page.evaluate(()=>window.viewer.bridgeVersion),2);
  await app.evaluate(({dialog},target)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:target});},saved);

  const initial=await width();
  const handle=await page.locator('.ag-header-cell[col-id="1"] .ag-header-cell-resize').boundingBox();
  await page.mouse.move(handle.x+handle.width/2,handle.y+handle.height/2);await page.mouse.down();
  await page.mouse.move(handle.x+handle.width/2+80,handle.y+handle.height/2,{steps:8});await page.mouse.up();
  await waitWidth(initial+80);
  await page.getByRole('tab',{name:'Other',exact:true}).click();
  await page.getByRole('tab',{name:'Data',exact:true}).click();await waitWidth(initial+80);
  await undo();await waitWidth(initial);await redo();await waitWidth(initial+80);
  console.log('PASS column drag persists across sheet switching and undo/redo');

  const adjacent=()=>page.locator('.ag-header-cell[col-id="2"]').evaluate(el=>el.getBoundingClientRect().width);
  const adjacentBefore=await adjacent();
  const shiftHandle=await page.locator('.ag-header-cell[col-id="1"] .ag-header-cell-resize').boundingBox();
  await page.keyboard.down('Shift');
  await page.mouse.move(shiftHandle.x+shiftHandle.width/2,shiftHandle.y+shiftHandle.height/2);await page.mouse.down();
  await page.mouse.move(shiftHandle.x+shiftHandle.width/2+35,shiftHandle.y+shiftHandle.height/2,{steps:5});await page.mouse.up();await page.keyboard.up('Shift');
  await waitWidth(initial+115);assert.ok(Math.abs(await adjacent()-(adjacentBefore-35))<2);
  await undo();await waitWidth(initial+80);assert.ok(Math.abs(await adjacent()-adjacentBefore)<2);
  console.log('PASS Shift-resize records both columns as one undo step');

  const rowHandle=await page.getByRole('separator',{name:'Resize row 2',exact:true}).boundingBox();
  const initialHeight=await rowHeight();
  await page.mouse.move(rowHandle.x+rowHandle.width/2,rowHandle.y+rowHandle.height/2);await page.mouse.down();
  await page.mouse.move(rowHandle.x+rowHandle.width/2,rowHandle.y+rowHandle.height/2+35,{steps:7});await page.mouse.up();
  await waitHeight(initialHeight+35);await undo();await waitHeight(initialHeight);await redo();await waitHeight(initialHeight+35);
  console.log('PASS row drag and undo/redo update rendered heights');

  await cell('2','0').click({button:'right'});
  await page.getByRole('menuitem',{name:/^Row height/}).click();
  await page.getByRole('dialog').getByRole('textbox',{name:'Height'}).fill('88');
  await page.getByRole('button',{name:'Apply',exact:true}).click();await waitHeight(88);
  await undo();await waitHeight(initialHeight+35);await redo();await waitHeight(88);
  console.log('PASS right-click row-height dialog updates grid immediately');

  await page.getByRole('combobox',{name:'Zoom',exact:true}).click();await page.getByRole('option',{name:'150%',exact:true}).click();
  await waitWidth((initial+80)*1.5);await waitHeight(132);
  await page.getByRole('combobox',{name:'Zoom',exact:true}).click();await page.getByRole('option',{name:'100%',exact:true}).click();
  await waitWidth(initial+80);await waitHeight(88);
  console.log('PASS zoom scales custom dimensions without changing stored sizes');

  // Sort before deleting: the menu must address the underlying spreadsheet row.
  await page.locator('.ag-header-cell[col-id="1"] .ag-header-cell-text').click();
  await page.getByText('Bob',{exact:true}).click({button:'right'});
  await page.getByRole('menuitem',{name:'Delete row(s)',exact:true}).click();
  await page.getByText('Bob',{exact:true}).waitFor({state:'detached'});
  await page.getByText('Carol',{exact:true}).waitFor();
  await undo();await page.getByText('Bob',{exact:true}).waitFor();
  console.log('PASS right-click delete targets correct row after sorting; undo restores it');

  await page.getByRole('button',{name:'Reset filters',exact:true}).click();
  // Disjoint checkbox selection must not remove the rows between the selections.
  await page.locator('.ag-row[row-id="2"] .ag-selection-checkbox input').check();
  await page.locator('.ag-row[row-id="4"] .ag-selection-checkbox input').check();
  await page.getByText('Alice',{exact:true}).click({button:'right'});
  await page.getByRole('menuitem',{name:'Delete row(s)',exact:true}).click();
  await page.getByText('Alice',{exact:true}).waitFor({state:'detached'});await page.getByText('Carol',{exact:true}).waitFor({state:'detached'});
  await page.getByText('Bob',{exact:true}).waitFor();await page.getByText('Dave',{exact:true}).waitFor();
  await undo();await page.getByText('Alice',{exact:true}).waitFor();await page.getByText('Carol',{exact:true}).waitFor();
  console.log('PASS disjoint selected-row delete is one undoable operation');

  await page.getByRole('button',{name:'Save',exact:true}).click();
  await page.locator('footer',{hasText:'All changes saved'}).waitFor();
  const persisted=XLSX.readFile(saved,{cellStyles:true});
  assert.equal(persisted.Sheets.Data.A2.v,'Alice');assert.equal(persisted.Sheets.Data.A4.v,'Carol');
  assert.ok(Math.abs(persisted.Sheets.Data['!cols'][1].width*7-(initial+80))<2);
  assert.ok(Math.abs(persisted.Sheets.Data['!rows'][1].hpx-88)<2);
  await page.getByRole('button',{name:'Save As',exact:true}).click();await page.locator('footer',{hasText:'All changes saved'}).waitFor();
  await app.evaluate(({dialog},target)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[target]});},saved);
  await page.getByRole('button',{name:'Open file',exact:true}).click();
  await page.getByText('Alice',{exact:true}).waitFor();await waitWidth(initial+80);await waitHeight(88);
  assert.deepEqual(errors,[]);
  await page.screenshot({path:path.join(out,packaged?'interactions-packaged.png':'interactions.png')});
  console.log(`PASS ${packaged?'packaged':'development'} Save / Save As bridge and persisted row/column sizes`);
 } catch(error) {
  console.error('Interaction failure:',error);
  await page.screenshot({path:path.join(out,'interaction-failure.png')}).catch(()=>{});
  throw error;
 } finally {
  await app.evaluate(({dialog})=>{dialog.showMessageBoxSync=()=>1;}).catch(()=>{});
  await app.close();
 }
}
main().catch(e=>{console.error(e);process.exitCode=1;});

