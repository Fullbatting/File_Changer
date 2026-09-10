// main.js
// Electron 메인 프로세스: 앱 창 생성 및 파일 다이얼로그(open/save) IPC 처리

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 640,
    backgroundColor: '#f4f6f9',
    icon: path.join(__dirname, 'assets', 'icon.png'),
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: false,
      spellcheck: false,
    },
  });

  mainWindow.setMenuBarVisibility(false);
  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
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

/**
 * 파일 열기 다이얼로그
 * args: { title, filters: [{name, extensions}], multi: boolean }
 * 반환: 선택된 파일 경로 배열 (취소 시 빈 배열)
 */
ipcMain.handle('dialog:openFile', async (event, args = {}) => {
  const { title = '파일 선택', filters = [{ name: 'All Files', extensions: ['*'] }], multi = false } = args;

  const properties = ['openFile'];
  if (multi) properties.push('multiSelections');

  const result = await dialog.showOpenDialog(mainWindow, {
    title,
    filters,
    properties,
  });

  if (result.canceled) return [];
  return result.filePaths;
});

/**
 * 파일 저장 다이얼로그
 * args: { title, defaultPath, filters: [{name, extensions}] }
 * 반환: 선택된 저장 경로 (취소 시 null)
 */
ipcMain.handle('dialog:saveFile', async (event, args = {}) => {
  const { title = '파일 저장', defaultPath = '', filters = [{ name: 'All Files', extensions: ['*'] }] } = args;

  const result = await dialog.showSaveDialog(mainWindow, {
    title,
    defaultPath,
    filters,
  });

  if (result.canceled) return null;
  return result.filePath;
});
