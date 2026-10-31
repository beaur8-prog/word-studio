'use strict';

const { contextBridge, ipcRenderer } = require('electron');
const MENU_ACTIONS = new Set(['new', 'open', 'save', 'save-as', 'export-docx', 'export-hwp', 'export-hwpx', 'export-pdf', 'print', 'updates', 'updates-check']);

// One operation across all document panels, including HWP/PDF exports.
// Keep the synchronous latch in the preload so rapid clicks cannot race IPC.
let fileOperation = null;
const fileOperationListeners = new Set();
function publishFileOperation() {
  for (const listener of fileOperationListeners) {
    try { listener(fileOperation); } catch { /* UI observers cannot strand a panel. */ }
  }
}
async function runFileOperation(kind, channel, ...args) {
  if (fileOperation) throw new Error('다른 파일 작업 중입니다. 열려 있는 선택창에서 선택하거나 취소해 주세요.');
  fileOperation = kind; publishFileOperation();
  try { return await ipcRenderer.invoke(channel, ...args); }
  finally { fileOperation = null; publishFileOperation(); }
}

contextBridge.exposeInMainWorld('wordDesktop', Object.freeze({
  platform: process.platform,
  getUpdateState: () => ipcRenderer.invoke('word-studio:get-update-state'),
  updateAction: (action, value) => {
    if (!['check', 'download', 'install', 'cancel', 'automatic', 'release-page'].includes(action) ||
        (action === 'automatic' ? typeof value !== 'boolean' : value !== undefined))
      throw new TypeError('Choose a supported update action.');
    return ipcRenderer.invoke('word-studio:update-action', action, value);
  },
  onUpdateState: callback => {
    if (typeof callback !== 'function') throw new TypeError('An update callback is required.');
    const listener = (_event, state) => callback(state);
    ipcRenderer.on('word-studio:update-state', listener);
    return () => ipcRenderer.removeListener('word-studio:update-state', listener);
  },
  outlineFonts: () => ipcRenderer.invoke('word-studio:outline-fonts'),
  readRecovery: () => ipcRenderer.invoke('word-studio:read-recovery'),
  saveRecovery: serialized => ipcRenderer.invoke('word-studio:save-recovery', serialized),
  saveFile: payload => runFileOperation('save', 'word-studio:save-file', payload),
  openFile: () => runFileOperation('open', 'word-studio:open-file'),
  onFileOperation: callback => {
    if (typeof callback !== 'function') throw new TypeError('A file-operation callback is required.');
    fileOperationListeners.add(callback); callback(fileOperation);
    return () => fileOperationListeners.delete(callback);
  },
  takeRequestedFile: () => ipcRenderer.invoke('word-studio:take-requested-file'),
  onOpenRequested: callback => {
    if (typeof callback !== 'function') throw new TypeError('A file-open callback is required.');
    const listener = () => callback();
    ipcRenderer.on('word-studio:open-requested', listener);
    return () => ipcRenderer.removeListener('word-studio:open-requested', listener);
  },
  exportPdf: payload => runFileOperation('pdf', 'word-studio:export-pdf', payload),
  print: () => runFileOperation('print', 'word-studio:print'),
  paste: () => ipcRenderer.invoke('word-studio:paste'),
  copyText: text => ipcRenderer.invoke('word-studio:copy-text', text),
  setTheme: theme => {
    if (theme !== 'dark' && theme !== 'light') throw new TypeError('Choose a dark or light theme.');
    ipcRenderer.send('word-studio:set-theme', theme);
  },
  onMenu: callback => {
    if (typeof callback !== 'function') throw new TypeError('A menu callback is required.');
    const listener = (_event, action) => { if (MENU_ACTIONS.has(action)) callback(action); };
    ipcRenderer.on('word-studio:menu', listener);
    return () => ipcRenderer.removeListener('word-studio:menu', listener);
  },
}));
