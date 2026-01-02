const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

function createWindow() {
    const win = new BrowserWindow({
        width: 1280,
        height: 800,
        icon: path.join(__dirname, 'assets', 'icon.png'),
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    win.maximize(); // Iniciar maximizado
    win.loadFile('index.html');
    // win.webContents.openDevTools(); // Descomentar para depurar
}

app.whenReady().then(() => {
    // Handler para obtener contactos de Supabase
    ipcMain.handle('get-supabase-contacts', async () => {
        try {
            const { getSupabaseContacts } = require('./supabaseHelper');
            const contacts = await getSupabaseContacts();
            console.log(`Supabase: ${contacts.length} contactos obtenidos.`);
            return { success: true, data: contacts };
        } catch (err) {
            console.error('Error en get-supabase-contacts:', err);
            return { success: false, error: err.message };
        }
    });

    // Handler para leer archivos de OneDrive
    ipcMain.handle('get-onedrive-files', async () => {
        try {
            // Construir ruta dinámica: C:\Users\USUARIO\OneDrive\DEVELOPMENT\UPTODATE
            const targetDir = path.join(os.homedir(), 'OneDrive', 'DEVELOPMENT', 'UPTODATE');

            console.log('Leyendo directorio:', targetDir);

            if (!fs.existsSync(targetDir)) {
                throw new Error(`La carpeta no existe: ${targetDir}`);
            }

            const result = [];

            // 1. Leer archivos de la raíz (UPTODATE)
            if (fs.existsSync(targetDir)) {
                const files = fs.readdirSync(targetDir);
                const excelFiles = files.filter(f => /\.(xlsx|xls|xlsm)$/i.test(f));

                for (const file of excelFiles) {
                    const fullPath = path.join(targetDir, file);
                    try {
                        const buffer = fs.readFileSync(fullPath);
                        result.push({
                            name: file,
                            data: buffer
                        });
                    } catch (readErr) {
                        console.error(`Error leyendo archivo raíz ${file}:`, readErr);
                    }
                }
            }

            // 2. Leer archivos de subcarpeta 'whatsapp'
            const whatsappDir = path.join(targetDir, 'whatsapp');
            if (fs.existsSync(whatsappDir)) {
                console.log('Leyendo subdirectorio whatsapp:', whatsappDir);
                const waFiles = fs.readdirSync(whatsappDir);
                const waExcel = waFiles.filter(f => /\.(xlsx|xls|xlsm)$/i.test(f));

                for (const file of waExcel) {
                    const fullPath = path.join(whatsappDir, file);
                    try {
                        const buffer = fs.readFileSync(fullPath);
                        // Opcional: ¿Queremos indicar que viene de whatsapp en el nombre? 
                        // El usuario no lo pidió explicitamente, pero ayuda a evitar colisiones.
                        // Dejaremos el nombre original, el app.js usa fileTitle.
                        result.push({
                            name: file,
                            data: buffer
                        });
                    } catch (readErr) {
                        console.error(`Error leyendo archivo whatsapp ${file}:`, readErr);
                    }
                }
            }

            return { success: true, files: result, count: result.length, path: targetDir };
        } catch (error) {
            console.error('Error en get-onedrive-files:', error);
            return { success: false, error: error.message };
        }
    });

    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

