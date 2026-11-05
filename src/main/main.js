// main.js
// Electron 메인 프로세스: 앱 창 만들기와 IPC(파일·폴더 대화상자, 사용자 데이터 폴더, 화면 준비 완료 알림).
// 시작 알림·오류 창은 startup.js, 창·IPC 보안은 app-security.js.

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { webPreferences, protectWindow, assertTrustedSender } = require('./app-security');
const { reportLaunch, watchProcess, watchWindow } = require('./startup');

watchProcess();

let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 640,
    backgroundColor: '#f4f6f9',
    webPreferences: webPreferences(),
  });

  protectWindow(mainWindow);
  watchWindow(mainWindow);
  mainWindow.setMenuBarVisibility(false);
  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

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
  assertTrustedSender(event, mainWindow);
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
  assertTrustedSender(event, mainWindow);
  const { title = '파일 저장', defaultPath = '', filters = [{ name: 'All Files', extensions: ['*'] }] } = args;

  const result = await dialog.showSaveDialog(mainWindow, {
    title,
    defaultPath,
    filters,
  });

  if (result.canceled) return null;
  return result.filePath;
});

/**
 * 앱 사용자 데이터 폴더 경로(서식 보관함·설정 저장)
 * Windows: %APPDATA%/doc-convert-mailmerge, macOS: ~/Library/Application Support/doc-convert-mailmerge
 * (예전 이름의 폴더를 그대로 써서 등록해 둔 서식·설정을 이어 쓴다)
 */
ipcMain.handle('app:getUserDataPath', (event) => {
  assertTrustedSender(event, mainWindow);
  return app.getPath('userData');
});

/**
 * 폴더 선택 다이얼로그(서식 폴더째 추가, 채운 서식 저장 폴더)
 * args: { title, defaultPath }
 * 반환: 선택된 폴더 경로 (취소 시 null)
 */
ipcMain.handle('dialog:selectFolder', async (event, args = {}) => {
  assertTrustedSender(event, mainWindow);
  const { title = '폴더 선택', defaultPath = '' } = args;

  const result = await dialog.showOpenDialog(mainWindow, {
    title,
    defaultPath,
    properties: ['openDirectory', 'createDirectory'],
  });

  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

/**
 * 화면이 시작을 마쳤다(서식 보관함·설정까지 열림) 또는 실패했다: 시작 프로그램에 알려 run.bat의 검은 창을 닫게 한다.
 * result: { ok: true } | { ok: false, error } — 실패면 화면이 이미 오류 창을 띄웠다
 */
ipcMain.handle('app:ready', (event, result = {}) => {
  assertTrustedSender(event, mainWindow);
  reportLaunch(result && result.ok === false ? { ok: false, error: `프로그램을 시작하지 못했습니다: ${String(result.error || '')}`, shown: true } : { ok: true });
  return true;
});
