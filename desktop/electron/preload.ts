// Copyright (c) 2026 AbeLin · MIT License
import { contextBridge, ipcRenderer, webUtils } from 'electron';
import type { CodeEvent, DesktopAPI, JobEvent } from '../shared/types';

const api: DesktopAPI = {
  health: () => ipcRenderer.invoke('python:health'),
  chooseSource: (kind) => ipcRenderer.invoke('dialog:source', kind),
  chooseOutput: () => ipcRenderer.invoke('dialog:output'),
  inspect: (path, headerRow) => ipcRenderer.invoke('excel:inspect', { path, headerRow }),
  preview: (path, sheet, headerRow) => ipcRenderer.invoke('excel:preview', { path, sheet, headerRow }),
  privacyInspect: (request) => ipcRenderer.invoke('excel:privacy_inspect', request),
  privacyExport: (request) => ipcRenderer.invoke('privacy:export', request),
  groups: (request) => ipcRenderer.invoke('excel:groups', request),
  start: (request) => ipcRenderer.invoke('job:start', request),
  cancel: (jobId) => ipcRenderer.invoke('job:cancel', jobId),
  onJob: (callback) => {
    const listener = (_event: Electron.IpcRendererEvent, data: JobEvent) => callback(data);
    ipcRenderer.on('job:event', listener);
    return () => ipcRenderer.removeListener('job:event', listener);
  },
  openOutput: (path) => ipcRenderer.invoke('output:open', path),
  history: () => ipcRenderer.invoke('history:list'),
  codeEnvironment: () => ipcRenderer.invoke('code:environment'),
  startCode: (request) => ipcRenderer.invoke('code:start', request),
  cancelCode: (jobId) => ipcRenderer.invoke('code:cancel', jobId),
  importScript: () => ipcRenderer.invoke('code:import'),
  saveText: (kind, text) => ipcRenderer.invoke('code:save-text', kind, text),
  copyText: (text) => ipcRenderer.invoke('code:copy-text', text),
  scripts: () => ipcRenderer.invoke('scripts:list'),
  script: (id) => ipcRenderer.invoke('scripts:get', id),
  saveScript: (request) => ipcRenderer.invoke('scripts:save', request),
  updateScript: (id, changes) => ipcRenderer.invoke('scripts:update', id, changes),
  exportScripts: () => ipcRenderer.invoke('scripts:export'),
  importScripts: () => ipcRenderer.invoke('scripts:import'),
  onCode: (callback) => {
    const listener = (_event: Electron.IpcRendererEvent, data: CodeEvent) => callback(data);
    ipcRenderer.on('code:event', listener);
    return () => ipcRenderer.removeListener('code:event', listener);
  },
  pathForFile: (file) => webUtils.getPathForFile(file),
  window: (action) => ipcRenderer.send('window:action', action),
};
contextBridge.exposeInMainWorld('desktop', api);
