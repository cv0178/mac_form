import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('desktopApi', {
  getAppDataPaths: () => ipcRenderer.invoke('desktop:get-app-data-paths'),
  selectKnowledgeFiles: () => ipcRenderer.invoke('desktop:select-knowledge-files'),
  selectFormFile: () => ipcRenderer.invoke('desktop:select-form-file'),
  selectDirectory: () => ipcRenderer.invoke('desktop:select-directory'),
  showOutputInFolder: (filePath) => ipcRenderer.invoke('desktop:show-output-in-folder', filePath),
  openOutputFile: (filePath) => ipcRenderer.invoke('desktop:open-output-file', filePath),
  notifyExportComplete: (outputName) => ipcRenderer.invoke('desktop:notify-export-complete', outputName),
});
