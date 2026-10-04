import {contextBridge,ipcRenderer,webUtils} from 'electron';
import type {ViewerAPI} from './types';
const api:ViewerAPI={
  bridgeVersion:2,
  open:()=>ipcRenderer.invoke('viewer:open'),openPath:path=>ipcRenderer.invoke('viewer:path',path),
  drop:file=>ipcRenderer.invoke('viewer:path',webUtils.getPathForFile(file)),
  sheet:name=>ipcRenderer.invoke('viewer:sheet',name),recent:()=>ipcRenderer.invoke('viewer:recent'),
  clearRecent:()=>ipcRenderer.invoke('viewer:clear'),copy:text=>ipcRenderer.invoke('viewer:copy',text),
  paste:()=>ipcRenderer.invoke('viewer:paste'),
  onOpen:callback=>{const listener=()=>callback();ipcRenderer.on('viewer:request-open',listener);return()=>ipcRenderer.removeListener('viewer:request-open',listener);},
  getPath:()=>ipcRenderer.invoke('viewer:getPath'),
  readBytes:path=>ipcRenderer.invoke('viewer:bytes',path),
  save:(base64,targetPath,format)=>ipcRenderer.invoke('viewer:save',base64,targetPath,format),
  chooseSavePath:suggestedName=>ipcRenderer.invoke('viewer:chooseSavePath',suggestedName),
};
contextBridge.exposeInMainWorld('viewer',api);
