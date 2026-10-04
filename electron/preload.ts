import {contextBridge,ipcRenderer,webUtils} from 'electron';
import type {ViewerAPI} from './types';
const api:ViewerAPI={
 open:()=>ipcRenderer.invoke('viewer:open'),openPath:path=>ipcRenderer.invoke('viewer:path',path),
 drop:file=>ipcRenderer.invoke('viewer:path',webUtils.getPathForFile(file)),
 sheet:name=>ipcRenderer.invoke('viewer:sheet',name),recent:()=>ipcRenderer.invoke('viewer:recent'),
 clearRecent:()=>ipcRenderer.invoke('viewer:clear'),copy:text=>ipcRenderer.invoke('viewer:copy',text),
 onOpen:callback=>{const listener=()=>callback();ipcRenderer.on('viewer:request-open',listener);return()=>ipcRenderer.removeListener('viewer:request-open',listener);}
};
contextBridge.exposeInMainWorld('viewer',api);
