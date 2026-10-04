import {app,BrowserWindow,ipcMain,dialog,Menu,clipboard} from 'electron';
import {Worker} from 'node:worker_threads';
import {join,extname,basename} from 'node:path';
import {pathToFileURL} from 'node:url';
import {stat,readFile,writeFile} from 'node:fs/promises';
import type {Book,Result} from './types';
const extensions=['xlsx','xls','xlsm','xlsb','csv'];
const dev=process.argv.includes('--dev')&&!app.isPackaged;
let win:BrowserWindow;
let worker:Worker|undefined;
let sequence=0;
let pendingPath=process.argv.find(p=>extensions.includes(extname(p).slice(1).toLowerCase()));
const entry=dev?'http://127.0.0.1:5173/':pathToFileURL(join(__dirname,'../dist/index.html')).href;
const recentFile=()=>join(app.getPath('userData'),'recent.json');
async function recent():Promise<string[]> {try {const data=JSON.parse(await readFile(recentFile(),'utf8'));return Array.isArray(data)?data.filter(p=>typeof p==='string').slice(0,10):[];}catch{return [];}}
function request(target:Worker,action:string,args:Record<string,unknown>={}):Promise<any> {
 return new Promise((resolve,reject)=>{
  const id=++sequence;
  const cleanup=()=>{clearTimeout(timer);target.off('message',message);target.off('error',error);target.off('exit',exited);};
  const error=(e:Error)=>{cleanup();reject(e);};
  const exited=()=>error(new Error('Spreadsheet reader stopped. Reopen the file.'));
  const message=(msg:any)=>{if(msg.id!==id)return;cleanup();msg.error?reject(new Error(msg.error)):resolve(msg.data);};
  const timer=setTimeout(()=>{error(new Error('Reading took too long. Try a smaller file.'));void target.terminate();},120000);
  target.on('message',message);target.once('error',error);target.once('exit',exited);
  target.postMessage({id,action,...args});
 });
}
async function openPath(path:unknown):Promise<Book> {
 if(typeof path!=='string'||!extensions.includes(extname(path).slice(1).toLowerCase()))throw new Error('Choose an xlsx, xls, xlsm, xlsb or csv file.');
 const info=await stat(path);
 if(!info.isFile())throw new Error('Please choose a file.');
 if(info.size>200*1024*1024)throw new Error('This MVP supports files up to 200 MB.');
 const next=new Worker(join(__dirname,'worker.js'),{resourceLimits:{maxOldGenerationSizeMb:1536}});
 // Keep an error listener even between requests to avoid a process-level crash.
 next.on('error',()=>{});
 try {
  const sheets=await request(next,'open',{path});
  const previous=worker;worker=next;void previous?.terminate();
  const paths=[path,...(await recent()).filter(p=>p!==path)].slice(0,10);
  await writeFile(recentFile(),JSON.stringify(paths)).catch(()=>{});
  app.addRecentDocument(path);
  return {name:basename(path),sheets,size:info.size};
 }catch(e){void next.terminate();throw e;}
}
let queue:Promise<unknown>=Promise.resolve();
function handle(channel:string,action:(...args:any[])=>unknown) {
 ipcMain.handle(channel,async(event,...args):Promise<Result<any>>=>{
  if(event.sender!==win.webContents||event.senderFrame!==win.webContents.mainFrame||event.senderFrame.url!==entry) return {ok:false,error:'Untrusted request.'};
  const task=queue.then(()=>action(...args));queue=task.catch(()=>{});
  try{return {ok:true,data:await task};}catch(e){return {ok:false,error:e instanceof Error?e.message:'Unexpected error.'};}
 });
}
if(!app.requestSingleInstanceLock())app.quit();
else {
 app.on('second-instance',(_event,argv)=>{pendingPath=argv.find(p=>extensions.includes(extname(p).slice(1).toLowerCase()));if(win){if(win.isMinimized())win.restore();win.focus();win.webContents.send('viewer:request-open');}});
 app.whenReady().then(()=>{
  win=new BrowserWindow({width:1280,height:820,minWidth:800,minHeight:500,show:false,backgroundColor:'#f7f9fb',webPreferences:{preload:join(__dirname,'preload.js'),contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}});
  win.webContents.setWindowOpenHandler(()=>({action:'deny'}));
  win.webContents.on('will-navigate',(e,url)=>{if(url!==entry)e.preventDefault();});
  win.webContents.session.setPermissionRequestHandler((_wc,_permission,callback)=>callback(false));
  handle('viewer:open',async()=>{
   if(pendingPath){const path=pendingPath;pendingPath=undefined;return openPath(path);}
   const result=await dialog.showOpenDialog(win,{properties:['openFile'],filters:[{name:'Spreadsheets',extensions}]});
   return result.canceled?null:openPath(result.filePaths[0]);
  });
  handle('viewer:path',openPath);
  handle('viewer:sheet',name=>{if(typeof name!=='string'||!worker)throw new Error('Open a workbook first.');return request(worker,'sheet',{name});});
  handle('viewer:recent',recent);
  handle('viewer:clear',async()=>{await writeFile(recentFile(),'[]');app.clearRecentDocuments();});
  handle('viewer:copy',text=>{if(typeof text!=='string'||text.length>10000000)throw new Error('Copy is limited to 10 million characters.');clipboard.writeText(text);});
  Menu.setApplicationMenu(Menu.buildFromTemplate([{label:'File',submenu:[{label:'Open…',accelerator:'CmdOrCtrl+O',click:()=>win.webContents.send('viewer:request-open')},{type:'separator'},{role:'quit'}]},{label:'View',submenu:[{role:'togglefullscreen'},...(dev?[{role:'toggleDevTools' as const}]:[])]}]));
  win.once('ready-to-show',()=>win.show());
  win.webContents.once('did-finish-load',()=>{if(pendingPath)win.webContents.send('viewer:request-open');});
  void win.loadURL(entry);
 });
 app.on('window-all-closed',()=>app.quit());
 app.on('before-quit',()=>{void worker?.terminate();});
}
