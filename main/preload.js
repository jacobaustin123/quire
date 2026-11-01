// Narrow bridge between the sandboxed renderer and the main process.
'use strict';

const { contextBridge, ipcRenderer, webUtils } = require('electron');

// Theme is applied here, before the page paints, so a forced Light/Dark never flashes the other.
function applyTheme(theme) {
  const el = document.documentElement;
  if (!el) return false;
  if (theme === 'light' || theme === 'dark') el.dataset.theme = theme;
  else delete el.dataset.theme;
  return true;
}
const initialTheme = ipcRenderer.sendSync('theme:get');
if (!applyTheme(initialTheme)) document.addEventListener('readystatechange', () => applyTheme(initialTheme), { once: true });
ipcRenderer.on('theme', (_e, theme) => applyTheme(theme));

contextBridge.exposeInMainWorld('quire', {
  init: () => ipcRenderer.invoke('doc:init'),
  chooseSavePath: (defaultName, format) => ipcRenderer.invoke('doc:choose-save-path', { defaultName, format }),
  write: (filePath, data) => ipcRenderer.invoke('doc:write', { filePath, data }),
  setDirty: (dirty) => ipcRenderer.send('doc:dirty', dirty),
  markPristine: () => ipcRenderer.send('doc:pristine'),
  confirm: (message, detail, ok) => ipcRenderer.invoke('confirm', { message, detail, ok }),
  onReplaceWord: (fn) => ipcRenderer.on('replace-word', (_e, info) => fn(info)),
  askRestore: (name, savedAt, fileChangedSince) => ipcRenderer.invoke('recovery:ask', { name, savedAt, fileChangedSince }),
  discardRecovery: () => ipcRenderer.send('recovery:discard'),
  reload: () => ipcRenderer.invoke('doc:reload'),
  onFileChanged: (fn) => ipcRenderer.on('file-changed', (_e, info) => fn(info)),
  onSaveRequest: (fn) => ipcRenderer.on('save-request', (_e, id) => fn(id)),
  saveResult: (id, ok) => ipcRenderer.send('save-result', { id, ok }),
  writeRecovery: (json) => ipcRenderer.send('recovery:write', json),
  clearRecovery: () => ipcRenderer.send('recovery:clear'),
  exportPdf: (defaultName) => ipcRenderer.invoke('pdf:export', { defaultName }),
  print: () => ipcRenderer.send('print'),
  openExternal: (url) => ipcRenderer.send('open-external', url),
  openPath: (p) => ipcRenderer.send('open-path', p),
  pathForFile: (file) => webUtils.getPathForFile(file),
  toggleTheme: () => ipcRenderer.send('toggle-theme'),
  onCommand: (fn) => ipcRenderer.on('cmd', (_e, name) => fn(name)),
  onMenu: (fn) => ipcRenderer.on('menu', (_e, name) => fn(name)),
});
