const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('pomodoroAPI', {
  showNotification: (title, body) => {
    return ipcRenderer.invoke('show-notification', { title, body });
  }
});
