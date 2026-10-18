'use strict';

const { contextBridge, ipcRenderer } = require('electron');
const MENU_ACTIONS = new Set(['new', 'open', 'save', 'save-as', 'export-docx', 'export-pdf', 'print']);

contextBridge.exposeInMainWorld('wordDesktop', Object.freeze({
  platform: process.platform,
  outlineFonts: () => ipcRenderer.invoke('word-studio:outline-fonts'),
  readRecovery: () => ipcRenderer.invoke('word-studio:read-recovery'),
  saveRecovery: serialized => ipcRenderer.invoke('word-studio:save-recovery', serialized),
  saveFile: payload => ipcRenderer.invoke('word-studio:save-file', payload),
  openFile: () => ipcRenderer.invoke('word-studio:open-file'),
  exportPdf: payload => ipcRenderer.invoke('word-studio:export-pdf', payload),
  print: () => ipcRenderer.invoke('word-studio:print'),
  paste: () => ipcRenderer.invoke('word-studio:paste'),
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
