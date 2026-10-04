// Editor end-to-end: open -> edit -> recalc -> undo/redo -> clipboard ->
// formatting -> sheet ops -> freeze/find -> save -> reopen-verify.
// Run from the repo root after `npm run build`:
//   node tests/editor.cjs
// Uses only repo-relative requires. Windows-only helper (taskkill) for orphan cleanup.
const { _electron: electron } = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execSync } = require('node:child_process');
const XLSX = require('xlsx');
const LOG = path.join(__dirname, 'editor-e2e.log');
const log = (m) => { try { fs.appendFileSync(LOG, new Date().toISOString() + ' ' + m + '\n'); } catch { } console.log(m); };

async function main() {
  try { fs.unlinkSync(LOG); } catch { }
  await new Promise((r) => setTimeout(r, 3000));
  const out = path.resolve('test-output');
  fs.mkdirSync(out, { recursive: true });
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([['Name', 'Amount', 'Total'], ['Alice', 12, 24], ['Bob', 8, 16], ['Carol', 30, 60]]);
  sheet.C2.f = 'B2*2'; sheet.C3.f = 'B3*2'; sheet.C4.f = 'B4*2';
  XLSX.utils.book_append_sheet(book, sheet, 'Sales');
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Second sheet'], ['Lazy loaded']]), 'Notes');
  const file = path.join(out, 'edit-me.xlsx');
  const saved = path.join(out, 'edited-copy.xlsx');
  XLSX.writeFile(book, file);
  // verify + settle so AV/indexer finishes with the fresh file
  const chk = XLSX.readFile(file);
  assert.equal(chk.Sheets['Sales']['B2'].v, 12);
  await new Promise((r) => setTimeout(r, 3000));
  log('fixture ready');

  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE;
  env.CLEAR_SHEET_TEST_USER_DATA=path.join(out,'profile-editor.cjs-'+process.pid);fs.mkdirSync(env.CLEAR_SHEET_TEST_USER_DATA,{recursive:true});
  const app = await electron.launch({ args: [path.resolve('.'), file], env, timeout: 60000 });
  log('launched');
  const page = await app.firstWindow({ timeout: 60000 });
  log('window');
  page.on('dialog', async (d) => { log('DIALOG ' + d.type() + ' ' + JSON.stringify(d.message())); try { await d.dismiss(); } catch { } });
  page.on('pageerror', (e) => log('PAGEERROR: ' + e.message));
  const step = async (name, fn) => {
    try { await fn(); log('PASS ' + name); }
    catch (e) { log('FAIL-STEP ' + name + ' :: ' + String(e).slice(0, 300)); throw e; }
  };
  const addrOf = (r, c) => {
    const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    let s = '', n = c;
    for (n++; n; n = Math.floor((n - 1) / 26)) s = A[(n - 1) % 26] + s;
    return s + (r + 1);
  };
  // click a grid cell and wait until the app selection actually lands there
  const clickCell = async (r, c) => {
    await page.locator('.ag-row[row-index="' + r + '"] [col-id="' + c + '"]').click({ timeout: 15000 });
    await page.waitForFunction((a) => document.querySelector('.address')?.textContent === a, addrOf(r, c), { timeout: 15000 });
  };

  try {
    await page.getByText('Alice', { exact: true }).waitFor({ timeout: 60000 });
    log('alice visible');

    await step('1 grid edit', async () => {
      await clickCell(1, 1);
      await page.keyboard.type('20');
      await page.keyboard.press('Enter');
      await page.locator('header .MuiChip-label').waitFor({ timeout: 15000 });
    });

    await step('2 recalc', async () => {
      await clickCell(1, 2);
      assert.equal(await page.getByRole('textbox', { name: 'Formula bar' }).inputValue(), '=B2*2');
      assert.equal((await page.locator('.ag-row[row-index="1"] [col-id="2"]').innerText()).trim(), '40');
    });

    await step('3 formula bar', async () => {
      await clickCell(1, 3);
      await page.getByRole('textbox', { name: 'Formula bar' }).fill('=SUM(B2:B4)');
      await page.keyboard.press('Enter');
      await page.waitForFunction(() => document.querySelector('.ag-row[row-index="1"] [col-id="3"]')?.textContent?.trim() === '58', null, { timeout: 15000 });
    });

    await step('4 undo/redo', async () => {
      await page.getByRole('button', { name: 'Undo' }).click({ timeout: 15000 });
      await page.waitForFunction(() => document.querySelector('.ag-row[row-index="1"] [col-id="3"]')?.textContent?.trim() === '', null, { timeout: 15000 });
      await page.getByRole('button', { name: 'Redo' }).click({ timeout: 15000 });
      await page.waitForFunction(() => document.querySelector('.ag-row[row-index="1"] [col-id="3"]')?.textContent?.trim() === '58', null, { timeout: 15000 });
    });

    await step('5 clipboard', async () => {
      await clickCell(1, 1);
      await page.keyboard.press('Control+c');
      await clickCell(3, 1);
      await page.keyboard.press('Control+v');
      await page.waitForFunction(() => document.querySelector('.ag-row[row-index="3"] [col-id="1"]')?.textContent?.trim() === '20', null, { timeout: 15000 });
    });

    await step('5b multi-cell paste and undo', async () => {
      await app.evaluate(({clipboard})=>clipboard.writeText('001\t2\r\nhello\t=1+2'));
      await clickCell(1,5);await page.keyboard.press('Control+v');
      await page.waitForFunction(()=>document.querySelector('.ag-row[row-index="2"] [col-id="6"]')?.textContent?.trim()==='3');
      await page.getByRole('button',{name:'Undo',exact:true}).click();
      await page.waitForFunction(()=>document.querySelector('.ag-row[row-index="1"] [col-id="5"]')?.textContent?.trim()==='');
    });
    await step('5c structural row insert/delete and undo', async () => {
      await clickCell(1,1);
      await page.getByRole('button',{name:'Insert menu',exact:true}).click();
      await page.getByRole('menuitem',{name:'Insert row(s)',exact:true}).click();
      await page.waitForFunction(()=>document.querySelector('.ag-row[row-index="2"] [col-id="0"]')?.textContent?.trim()==='Alice');
      await page.getByRole('button',{name:'Undo',exact:true}).click();
      await page.waitForFunction(()=>document.querySelector('.ag-row[row-index="1"] [col-id="0"]')?.textContent?.trim()==='Alice');
    });
    await step('6 format', async () => {
      await clickCell(0, 0);
      await page.getByRole('button', { name: 'Bold' }).click({ timeout: 15000 });
      const w = await page.evaluate(() => getComputedStyle(document.querySelector('.ag-row[row-index="0"] [col-id="0"]')).fontWeight);
      assert.ok(w === '700' || w === 'bold', 'bold applied in grid, got ' + w);
      await clickCell(1, 1);
      await page.getByRole('combobox', { name: 'Number' }).click({ timeout: 15000 });
      await page.getByRole('option', { name: 'Currency ($)' }).click({ timeout: 15000 });
      await page.waitForFunction(() => document.querySelector('.ag-row[row-index="1"] [col-id="1"]')?.textContent?.trim() === '$20.00', null, { timeout: 15000 });
    });

    await step('7 sheets', async () => {
      await page.getByRole('button', { name: 'Add sheet', exact: true }).click({ timeout: 15000 });
      await page.getByRole('tab', { name: /Sheet/ }).first().waitFor({ timeout: 15000 });
      await page.getByRole('button', { name: 'Sheet menu' }).click({ timeout: 15000 });
      await page.getByRole('menuitem', { name: 'Rename…' }).click({ timeout: 15000 });
      await page.getByLabel('Sheet name').fill('Renamed');
      await page.getByRole('button', { name: 'Rename', exact: true }).click({ timeout: 15000 });
      await page.getByRole('tab', { name: 'Renamed' }).waitFor({ timeout: 15000 });
    });

    await step('8 freeze+find', async () => {
      await page.getByRole('tab', { name: 'Sales', exact: true }).click({ timeout: 15000 });
      await page.getByText('Alice', { exact: true }).waitFor({ timeout: 15000 });
      await page.getByRole('button', { name: 'Freeze menu' }).click({ timeout: 15000 });
      await page.getByRole('menuitem', { name: 'Freeze first row' }).click({ timeout: 15000 });
      await page.keyboard.press('Control+h');
      await page.getByRole('textbox', { name: 'Find' }).fill('Alice');
      await page.getByRole('button', { name: 'Find all' }).click({ timeout: 15000 });
      await page.getByText(/1 \/ 1 matches/).waitFor({ timeout: 15000 });
      await page.keyboard.press('Escape');
    });

    await step('8b close cancelled keeps unsaved workbook', async () => {
      await app.evaluate(({dialog,BrowserWindow})=>{globalThis.closePrompts=0;dialog.showMessageBoxSync=()=>{globalThis.closePrompts++;return 0;};BrowserWindow.getAllWindows()[0].close();});
      await page.waitForTimeout(250);
      assert.equal(page.isClosed(),false);
      assert.equal(await app.evaluate(()=>globalThis.closePrompts),1);
    });
    await step('8c cancelled and failed Save As retain dirty state and original', async () => {
      const before=fs.readFileSync(file);
      await app.evaluate(({dialog})=>{dialog.showSaveDialog=async()=>({canceled:true});});
      await page.getByRole('button',{name:'Save As',exact:true}).click();
      await page.getByRole('button',{name:'Save As',exact:true}).waitFor({state:'visible'});
      await page.waitForFunction(()=>!document.querySelector('main[aria-busy="true"]'));
      assert.ok(await page.locator('footer').innerText().then(t=>t.includes('Unsaved changes')));
      await app.evaluate(({dialog},target)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:target});},path.join(out,'missing-'+Date.now(),'fail.xlsx'));
      await page.getByRole('button',{name:'Save As',exact:true}).click();
      await page.getByRole('alert').filter({hasText:/ENOENT/}).waitFor();
      assert.ok((await page.locator('footer').innerText()).includes('Unsaved changes'));
      assert.deepEqual(fs.readFileSync(file),before);
    });

    await step('9 save', async () => {
      await app.evaluate(({dialog},saved)=>{dialog.showSaveDialog=async()=>({canceled:false,filePath:saved});},saved);
      await page.getByRole('button', { name: 'Save', exact: true }).click({ timeout: 15000 });
      await page.locator('footer', { hasText: 'All changes saved' }).waitFor({ timeout: 25000 });
    });
    log('all steps done; closing');
  } finally {
    try { await Promise.race([app.close(), new Promise((_, rej) => setTimeout(() => rej(new Error('close-timeout')), 20000))]); log('closed cleanly'); }
    catch (e) { log('close-hung; force killing'); app.process().kill(); }
  }

  assert.equal(XLSX.readFile(file).Sheets.Sales.B2.v,12,'original retained');
  const wb = XLSX.readFile(saved, { cellFormula: true, cellStyles: true });
  assert.equal(wb.Sheets['Sales']['B2'].v, 20);
  assert.equal(wb.Sheets['Sales']['C2'].f, 'B2*2');
  assert.equal(wb.Sheets['Sales']['D2'].f, 'SUM(B2:B4)');
  assert.equal(wb.Sheets['Sales']['B4'].v, 20);
  assert.equal(wb.Sheets['Sales']['B2'].z, '$#,##0.00');
  assert.ok(wb.SheetNames.includes('Renamed'), 'renamed sheet persisted: ' + wb.SheetNames);
  log('PASS reopen verify: ' + wb.SheetNames);
  log('ALL EDITOR E2E TESTS PASSED');
}
main().catch((e) => { log('FAIL ' + String(e && e.stack || e).slice(0, 600)); process.exitCode = 1; });
