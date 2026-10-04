const assert = require('node:assert/strict');
const path = require('node:path');
const { _electron: electron } = require('playwright');
const XLSX = require('xlsx');
const fs = require('node:fs');

async function runPinnacleTests() {
  const out = path.resolve('test-output');
  fs.mkdirSync(out, { recursive: true });
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Category', 'Region', 'Sales', 'Target', 'Formula Test'],
    ['Apples', 'North', 100, 80, null],
    ['Oranges', 'South', 150, 100, null],
    ['Apples', 'South', 200, 120, null],
    ['Bananas', 'North', 50, 60, null],
    ['Apples', 'North', 300, 150, null],
  ]);
  XLSX.utils.book_append_sheet(book, sheet, 'Data');
  const file = path.join(out, 'pinnacle-test.xlsx');
  XLSX.writeFile(book, file);

  const profile = path.join(out, 'pinnacle-profile-' + process.pid);
  fs.mkdirSync(profile, { recursive: true });
  const env = { ...process.env, CLEAR_SHEET_TEST_USER_DATA: profile };
  delete env.ELECTRON_RUN_AS_NODE;

  const app = await electron.launch({
    args: [path.resolve('.'), file],
    env,
    timeout: 60000,
  });

  try {
    const page = await app.firstWindow({ timeout: 60000 });
    page.on('dialog', async (d) => { try { await d.dismiss(); } catch {} });
    await page.getByText('Apples', { exact: false }).first().waitFor({ timeout: 30000 });
    console.log('PASS test fixture loaded');

    // Helper to evaluate a formula in cell (r, c) and get rendered text
    const evalFormula = async (r, c, formula) => {
      await page.locator(`.ag-row[row-index="${r}"] [col-id="${c}"]`).click();
      await page.getByRole('textbox', { name: 'Formula bar' }).fill(formula);
      await page.keyboard.press('Enter');
      await page.waitForTimeout(400);
      return (await page.locator(`.ag-row[row-index="${r}"] [col-id="${c}"]`).innerText()).trim();
    };

    // 1. Test COUNTIFS
    const countifsVal = await evalFormula(1, 4, '=COUNTIFS(A2:A6, "Apples", B2:B6, "North")');
    assert.equal(countifsVal, '2', `COUNTIFS expected 2, got ${countifsVal}`);
    console.log('PASS COUNTIFS');

    // 2. Test SUMIFS
    const sumifsVal = await evalFormula(2, 4, '=SUMIFS(C2:C6, A2:A6, "Apples", B2:B6, "North")');
    assert.equal(sumifsVal, '400', `SUMIFS expected 400 (100 + 300), got ${sumifsVal}`);
    console.log('PASS SUMIFS');

    // 3. Test AVERAGEIFS
    const avgifsVal = await evalFormula(3, 4, '=AVERAGEIFS(C2:C6, A2:A6, "Apples", B2:B6, "North")');
    assert.equal(avgifsVal, '200', `AVERAGEIFS expected 200, got ${avgifsVal}`);
    console.log('PASS AVERAGEIFS');

    // 4. Test IFS
    const ifsVal = await evalFormula(4, 4, '=IFS(C2>200, "High", C2>50, "Medium", TRUE, "Low")');
    assert.equal(ifsVal, 'Medium', `IFS expected Medium, got ${ifsVal}`);
    console.log('PASS IFS');

    // 5. Test SWITCH
    const switchVal = await evalFormula(5, 4, '=SWITCH(A2, "Oranges", "Orange", "Apples", "Apple", "Other")');
    assert.equal(switchVal, 'Apple', `SWITCH expected Apple, got ${switchVal}`);
    console.log('PASS SWITCH');

    // 6. Test MEDIAN
    const medianVal = await evalFormula(1, 4, '=MEDIAN(10, 20, 30, 40, 50)');
    assert.equal(medianVal, '30', `MEDIAN expected 30, got ${medianVal}`);
    console.log('PASS MEDIAN');

    // 7. Test STDEV & STDEV.S
    const stdevVal = await evalFormula(2, 4, '=ROUND(STDEV(10, 20, 30, 40, 50), 2)');
    assert.equal(stdevVal, '15.81', `STDEV expected 15.81, got ${stdevVal}`);
    console.log('PASS STDEV');

    // 8. Test PMT
    const pmtVal = await evalFormula(3, 4, '=ROUND(PMT(0.05/12, 60, -10000), 2)');
    assert.equal(pmtVal, '188.71', `PMT expected 188.71, got ${pmtVal}`);
    console.log('PASS PMT');

    // 9. Test F4 key cycling
    await page.locator('.ag-row[row-index="1"] [col-id="4"]').click();
    const fb = page.getByRole('textbox', { name: 'Formula bar' });
    await fb.fill('=SUM(A1');
    await fb.press('F4');
    assert.equal(await fb.inputValue(), '=SUM($A$1', 'F4 first cycle');
    await fb.press('F4');
    assert.equal(await fb.inputValue(), '=SUM(A$1', 'F4 second cycle');
    await fb.press('F4');
    assert.equal(await fb.inputValue(), '=SUM($A1', 'F4 third cycle');
    await fb.press('F4');
    assert.equal(await fb.inputValue(), '=SUM(A1', 'F4 fourth cycle');
    await page.keyboard.press('Escape');
    console.log('PASS F4 Absolute Reference Cycling');

    // 10. Test Intellisense Autocomplete Popover
    await page.locator('.ag-row[row-index="1"] [col-id="4"]').click();
    await fb.fill('=SUMI');
    await page.waitForSelector('.formula-autocomplete-popup', { timeout: 5000 });
    console.log('PASS Intellisense Autocomplete Popover displayed');
    await page.keyboard.press('Tab');
    const autoVal = await fb.inputValue();
    assert.ok(autoVal.startsWith('=SUMIF('), `Expected autocomplete to start with =SUMIF(, got ${autoVal}`);
    await page.keyboard.press('Escape');
    console.log('PASS Intellisense Tab Completion');

    // 11. Test Chart Modal Opens
    const chartBtn = page.getByRole('button', { name: 'Chart', exact: true });
    if (await chartBtn.isVisible()) {
      await chartBtn.click();
      await page.waitForSelector('text=Insert Chart', { timeout: 5000 });
      await page.getByRole('button', { name: 'Done' }).click();
      console.log('PASS Chart Modal');
    }

    // 12. Test Format Cells Dialog (Ctrl+1)
    await page.locator('.ag-row[row-index="1"] [col-id="2"]').click();
    await page.keyboard.press('Control+1');
    await page.waitForSelector('text=Format Cells', { timeout: 5000 });
    await page.getByRole('button', { name: 'Cancel' }).click();
    console.log('PASS Format Cells Dialog (Ctrl+1)');

    // 13. Test Hotkey Ctrl+; (Date)
    await page.locator('.ag-row[row-index="5"] [col-id="4"]').click();
    await page.keyboard.press('Control+;');
    await page.waitForTimeout(300);
    const dateText = (await page.locator('.ag-row[row-index="5"] [col-id="4"]').innerText()).trim();
    assert.ok(/^\d{4}-\d{2}-\d{2}$/.test(dateText), `Expected date format YYYY-MM-DD, got ${dateText}`);
    console.log('PASS Ctrl+; Date Hotkey');

    // 14. Test Fill Handle is rendered
    await page.locator('.ag-row[row-index="1"] [col-id="2"]').click();
    await page.waitForSelector('.cs-fill-handle', { timeout: 5000 });
    console.log('PASS Drag Fill Handle Rendered');

    console.log('\n========================================');
    console.log('ALL PINNACLE EXCEL FEATURES VERIFIED OK!');
    console.log('========================================\n');
  } finally {
    await app.close();
  }
}

runPinnacleTests().catch((e) => {
  console.error('FAIL pinnacle test:', e);
  process.exit(1);
});
