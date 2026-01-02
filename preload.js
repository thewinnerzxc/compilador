const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    getOneDriveFiles: () => ipcRenderer.invoke('get-onedrive-files'),
    getSupabaseContacts: () => ipcRenderer.invoke('get-supabase-contacts')
});
