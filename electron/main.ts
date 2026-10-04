import {app,BrowserWindow,ipcMain,dialog,Menu,clipboard} from 'electron';
import {Worker} from 'node:worker_threads';
import {join,extname,basename,dirname,resolve} from 'node:path';
import {pathToFileURL} from 'node:url';
import {stat,readFile,writeFile,rename,unlink,open} from 'node:fs/promises';
import {randomBytes} from 'node:crypto';
import type {Book,Result,Sheet} from './types';
const profile=app.commandLine.getSwitchValue('user-data-dir') || (!app.isPackaged ? process.env.CLEAR_SHEET_TEST_USER_DATA : undefined) || ((process.argv.includes('--dev')||app.commandLine.hasSwitch('dev'))&&!app.isPackaged ? join(app.getPath('appData'),'clear-sheet-dev') : undefined);
if(profile)app.setPath('userData',resolve(profile));
const extensions=['xlsx','xls','xlsm','xlsb','csv'];
const dev=(process.argv.includes('--dev')||app.commandLine.hasSwitch('dev'))&&!app.isPackaged;
let win:BrowserWindow;
let worker:Worker|undefined;
let sequence=0;
let pendingPath=process.argv.find(p=>extensions.includes(extname(p).slice(1).toLowerCase()));
let currentPath:string|undefined;
let openedPath:string|undefined;
const authorized = new Set<string>();
const pathKey=(p:string)=>resolve(p).toLowerCase();
const entry=dev?'http://127.0.0.1:5173/':pathToFileURL(join(__dirname,'../dist/index.html')).href;
const recentFile=()=>join(app.getPath('userData'),'recent.json');
async function recent():Promise<string[]> {try {const data=JSON.parse(await readFile(recentFile(),'utf8'));return Array.isArray(data)?data.filter(p=>typeof p==='string').slice(0,10):[];}catch{return [];}}
async function remember(path:string) {
  currentPath=path;
  const paths=[path,...(await recent()).filter(p=>p!==path)].slice(0,10);
  await writeFile(recentFile(),JSON.stringify(paths)).catch(()=>{});
  app.addRecentDocument(path);
}
function request<T>(target:Worker,action:string,args:Record<string,unknown>={}):Promise<T> {
  return new Promise((resolve,reject)=>{
    const id=++sequence;
    const cleanup=()=>{clearTimeout(timer);target.off('message',message);target.off('error',error);target.off('exit',exited);};
    const error=(e:Error)=>{cleanup();reject(e);};
    const exited=()=>error(new Error('Spreadsheet reader stopped. Reopen the file.'));
    const message=(msg:{id:number;data:T;error?:string})=>{if(msg.id!==id)return;cleanup();msg.error?reject(new Error(msg.error)):resolve(msg.data);};
    const timer=setTimeout(()=>{error(new Error('Reading took too long. Try a smaller file.'));void target.terminate();},120000);
    target.on('message',message);target.once('error',error);target.once('exit',exited);
    target.postMessage({id,action,...args});
  });
}
function validExt(path:string):boolean {return extensions.includes(extname(path).slice(1).toLowerCase());}
async function openPath(path:unknown):Promise<Book> {
  if(typeof path!=='string'||!validExt(path))throw new Error('Choose an xlsx, xls, xlsm, xlsb or csv file.');
  const info=await stat(path);
  if(!info.isFile())throw new Error('Please choose a file.');
  if(info.size>200*1024*1024)throw new Error('This MVP supports files up to 200 MB.');
  const next=new Worker(join(__dirname,'worker.js'),{resourceLimits:{maxOldGenerationSizeMb:1536}});
  // Keep an error listener even between requests to avoid a process-level crash.
  next.on('error',()=>{});
  try {
    const sheets=await request<string[]>(next,'open',{path});
    const previous=worker;worker=next;void previous?.terminate();
    openedPath=resolve(path);
    authorized.clear();authorized.add(pathKey(path));
    await remember(path);
    return {name:basename(path),sheets,size:info.size,path};
  }catch(e){void next.terminate();throw e;}
}
/** Atomically write buffer to target: tmp file in same dir + rename. Never truncates target on failure. */
async function atomicWrite(target:string, data:Buffer):Promise<void> {
  const dir=dirname(target);
  const tmp=join(dir,`.${basename(target)}.${process.pid}.${randomBytes(6).toString('hex')}.tmp`);
  try {
    const file = await open(tmp,'wx');
    try { await file.writeFile(data); await file.sync(); } finally { await file.close(); }
    await rename(tmp,target);
  }catch(e){await unlink(tmp).catch(()=>{});throw e;}
}
function validBase64(s:unknown):s is string {
  return typeof s==='string'&&s.length>0&&s.length<600*1024*1024&&/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(s);
}
let queue:Promise<unknown>=Promise.resolve();
function handle<Args extends unknown[],T>(channel:string,action:(...args:Args)=>T) {
  ipcMain.handle(channel,async(event,...args:Args):Promise<Result<Awaited<T>>>=>{
    if(event.sender!==win.webContents||event.senderFrame!==win.webContents.mainFrame||event.senderFrame.url!==entry) return {ok:false,error:'Untrusted request.'};
    const task=queue.then(()=>action(...args));queue=task.catch(()=>{});
    try{return {ok:true,data:await task};}catch(e){return {ok:false,error:e instanceof Error?e.message:'Unexpected error.'};}
  });
}
if(!dev && !app.requestSingleInstanceLock())app.quit();
else {
  app.on('second-instance',(_event,argv)=>{pendingPath=argv.find(p=>extensions.includes(extname(p).slice(1).toLowerCase()));if(win){if(win.isMinimized())win.restore();win.focus();win.webContents.send('viewer:request-open');}});
  app.whenReady().then(()=>{
    win=new BrowserWindow({width:1360,height:860,minWidth:800,minHeight:500,show:false,icon:app.isPackaged?join(process.resourcesPath,'icon.ico'):join(__dirname,'../build/icon.ico'),backgroundColor:'#f7f9fb',webPreferences:{preload:join(__dirname,'preload.js'),contextIsolation:true,nodeIntegration:false,sandbox:true,webSecurity:true}});
    win.webContents.on('page-title-updated',event=>{event.preventDefault();win.setTitle(`Clear Sheet ${app.getVersion()}`);});
    win.webContents.on('will-prevent-unload', event => {
      const choice=dialog.showMessageBoxSync(win,{type:'warning',buttons:['Keep editing','Discard and close'],defaultId:0,cancelId:0,title:'Unsaved changes',message:'Unsaved changes or a save in progress. Close without saving?'});
      if(choice===1)event.preventDefault();
    });
    win.webContents.setWindowOpenHandler(()=>({action:'deny'}));
    win.webContents.on('will-navigate',(e,url)=>{if(url!==entry)e.preventDefault();});
    win.webContents.session.setPermissionRequestHandler((_wc,_permission,callback)=>callback(false));
    handle('viewer:open',async()=>{
      if(pendingPath){const path=pendingPath;pendingPath=undefined;return openPath(path);}
      const result=await dialog.showOpenDialog(win,{properties:['openFile'],filters:[{name:'Spreadsheets',extensions}]});
      return result.canceled?null:openPath(result.filePaths[0]);
    });
    handle('viewer:path',openPath);
    handle('viewer:sheet',name=>{if(typeof name!=='string'||!worker)throw new Error('Open a workbook first.');return request<Sheet>(worker,'sheet',{name});});
    handle('viewer:recent',recent);
    handle('viewer:clear',async()=>{await writeFile(recentFile(),'[]');app.clearRecentDocuments();});
    handle('viewer:copy',text=>{if(typeof text!=='string'||text.length>10000000)throw new Error('Copy is limited to 10 million characters.');clipboard.writeText(text);});
    handle('viewer:paste',()=>clipboard.readText());
    handle('viewer:getPath',()=>currentPath??null);
    handle('viewer:bytes',async(targetPath?:unknown)=>{
      const p=typeof targetPath==='string'?targetPath:currentPath;
      if(!p||!validExt(p)||!authorized.has(pathKey(p)))throw new Error('Open a workbook first.');
      const info=await stat(p);
      if(!info.isFile())throw new Error('File no longer exists.');
      if(info.size>200*1024*1024)throw new Error('This MVP supports files up to 200 MB.');
      const buf=await readFile(p);
      return {base64:buf.toString('base64'),path:p};
    });
    handle('viewer:save',async(base64:unknown,targetPath:unknown,format:unknown)=>{
      if(!validBase64(base64))throw new Error('Nothing to save.');
      const p=typeof targetPath==='string'?targetPath:currentPath;
      if(!p||!validExt(p)||!authorized.has(pathKey(p)))throw new Error('Choose a file location first (Save As).');
      if(format!==extname(p).slice(1).toLowerCase())throw new Error('Output format does not match file extension.');
      if(openedPath && pathKey(p)===pathKey(openedPath) && extname(p).toLowerCase()!=='.csv')throw new Error('Save an edited copy to preserve the original Excel file.');
      const data=Buffer.from(base64 as string,'base64');
      if(!data.length)throw new Error('Nothing to save.');
      await atomicWrite(p,data);
      await remember(p);
      return {path:p};
    });
    handle('viewer:chooseSavePath',async(suggestedName:unknown)=>{
      const defName=typeof suggestedName==='string'&&suggestedName?suggestedName:'workbook.xlsx';
      const result=await dialog.showSaveDialog(win,{defaultPath:defName,filters:extensions.map(ext=>({name:ext.toUpperCase(),extensions:[ext]}))});
      if(result.canceled||!result.filePath)return null;
      let path=resolve(result.filePath);
      if(!extname(path))path+='.xlsx';
      if(!validExt(path))throw new Error('Unsupported file extension.');
      authorized.add(pathKey(path));
      return {path,format:extname(path).slice(1).toLowerCase()};
    });
    Menu.setApplicationMenu(Menu.buildFromTemplate([{label:'File',submenu:[{label:'Open…',accelerator:'CmdOrCtrl+O',click:()=>win.webContents.send('viewer:request-open')},{type:'separator'},{role:'quit'}]},{label:'View',submenu:[{role:'togglefullscreen'},...(dev?[{role:'toggleDevTools' as const}]:[])]}]));
    win.once('ready-to-show',()=>win.show());
    win.webContents.once('did-finish-load',()=>{if(pendingPath)win.webContents.send('viewer:request-open');});
    void win.loadURL(entry);
  });
  app.on('window-all-closed',()=>app.quit());
  app.on('will-quit',()=>{void worker?.terminate();});
}
