const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('aerizenDesktop', {
  platform: () => ipcRenderer.invoke('aerizen:platform'),
});
