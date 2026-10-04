const {_electron:electron}=require('playwright');
const fs=require('node:fs');
const path=require('node:path');
const assert=require('node:assert/strict');
async function main(){
 const out=path.resolve('test-output');fs.mkdirSync(out,{recursive:true});
 const profile=path.join(out,'packaged-profile-'+process.pid);fs.mkdirSync(profile,{recursive:true});
 const file=path.join(out,'packaged.csv');fs.writeFileSync(file,'Name,Value\r\nPackaged,4\r\n');
 const env={...process.env};delete env.ELECTRON_RUN_AS_NODE;
 const app=await electron.launch({executablePath:path.resolve('release/win-unpacked/Clear Sheet.exe'),args:['--user-data-dir='+profile,file],env,timeout:60000});
 try {
  const page=await app.firstWindow();
  await page.getByText('Packaged',{exact:true}).waitFor();
  await page.locator('.ag-row[row-index="1"] [col-id="1"]').click();
  await page.getByRole('textbox',{name:'Formula bar'}).fill('17');
  await page.keyboard.press('Enter');
  await page.getByRole('button',{name:'Save',exact:true}).click();
  await page.locator('footer',{hasText:'All changes saved'}).waitFor();
  assert.match(fs.readFileSync(file,'utf8'),/Packaged,17/);
  assert.equal(await page.evaluate(()=>typeof window.require),'undefined');
  console.log('PASS packaged Windows app: startup, preload, CSV edit/save and renderer isolation');
 }finally{await app.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});
