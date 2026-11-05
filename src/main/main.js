// main.js
// Electron 메인 프로세스: 앱 창 생성, 파일 다이얼로그(open/save) IPC 처리, 문제가 생기면 오류 창으로 알리기

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { webPreferences, protectWindow, assertTrustedSender } = require('./app-security');

let mainWindow = null;

/* ---------- 오류 창 ----------
 * 프로그램은 검은 창(cmd) 없이 실행되므로, 문제가 생기면 화면에 오류 창으로 알린다. 같은 내용은 잠깐 동안 한 번만. */
const ERROR_TITLE = '사회사업실 도우미 - 오류';
const recentErrors = new Map();
function showError(message, err) {
  const detail = err ? String((err && err.message) || err) : '';
  const key = `${message}|${detail}`;
  const now = Date.now();
  if (recentErrors.has(key) && now - recentErrors.get(key) < 5000) return;
  recentErrors.set(key, now);
  const text = detail ? `${message}\n\n${detail}` : message;
  try {
    if (app.isReady()) dialog.showErrorBox(ERROR_TITLE, text);
    else console.error(text);
  } catch (e) {
    console.error(text);
  }
}

process.on('uncaughtException', (err) => showError('예상하지 못한 오류가 발생했습니다. 같은 일이 되풀이되면 프로그램을 다시 실행하세요.', err));
process.on('unhandledRejection', (err) => showError('예상하지 못한 오류가 발생했습니다. 같은 일이 되풀이되면 프로그램을 다시 실행하세요.', err));

/** 화면 쪽 문제(준비 실패·멈춤·꺼짐·불러오기 실패)를 오류 창으로 */
function watchWindow(win) {
  const wc = win.webContents;
  wc.on('preload-error', (e, preloadPath, err) => showError('프로그램 화면을 준비하지 못했습니다. run.bat을 두 번 눌러 다시 준비한 뒤 실행하세요.', err));
  wc.on('did-fail-load', (e, code, desc, url, isMainFrame) => {
    if (isMainFrame && code !== -3) showError('프로그램 화면을 불러오지 못했습니다.', `${desc} (${code})`);
  });
  wc.on('render-process-gone', async (e, details) => {
    if (details.reason === 'clean-exit' || win.isDestroyed()) return;
    const r = await dialog.showMessageBox(win, {
      type: 'error',
      title: ERROR_TITLE,
      message: '프로그램 화면이 갑자기 꺼졌습니다.',
      detail: `입력하던 인적사항은 지워졌습니다. 다시 불러와 처음부터 입력하세요. (${details.reason}, 코드 ${details.exitCode})`,
      buttons: ['다시 불러오기', '프로그램 닫기'],
      defaultId: 0,
      noLink: true,
    });
    if (win.isDestroyed()) return;
    if (r.response === 0) wc.reload();
    else win.close();
  });
  win.on('unresponsive', async () => {
    const r = await dialog.showMessageBox(win, {
      type: 'warning',
      title: '사회사업실 도우미',
      message: '프로그램이 응답하지 않습니다.',
      detail: '한글로 큰 서식을 다루는 중이면 잠시 기다리면 돌아옵니다.',
      buttons: ['기다리기', '프로그램 닫기'],
      defaultId: 0,
      noLink: true,
    });
    if (r.response === 1 && !win.isDestroyed()) win.destroy();
  });
}

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
 * 앱 사용자 데이터 폴더 경로 반환 (동의어 사전 저장용)
 * Windows: %APPDATA%/doc-convert-mailmerge, macOS: ~/Library/Application Support/doc-convert-mailmerge
 * (예전 통합 프로그램과 같은 폴더 이름을 그대로 써서, 저장해 둔 동의어 사전을 이어 쓴다)
 */
ipcMain.handle('app:getUserDataPath', (event) => {
  assertTrustedSender(event, mainWindow);
  return app.getPath('userData');
});


/**
 * 폴더 선택 다이얼로그 (메일머지 일괄 생성 결과를 저장할 폴더 선택용)
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
