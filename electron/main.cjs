const { app, BrowserWindow, shell, ipcMain } = require('electron');
const path = require('path');

const isDev = !app.isPackaged;
app.setName('Aerizen Asset Management v5.5');

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1100,
    minHeight: 760,
    title: 'Aerizen Asset Management v5.5',
    icon: isDev ? path.join(__dirname, '..', 'public', 'favicon.png') : path.join(__dirname, '..', 'dist', 'favicon.png'),
    backgroundColor: '#0f172a',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  if (isDev) {
    win.loadURL('http://127.0.0.1:5173');
  } else {
    win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

app.whenReady().then(() => {
  ipcMain.handle('aerizen:platform', () => ({
    platform: process.platform,
    version: app.getVersion(),
    packaged: app.isPackaged,
  }));
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
