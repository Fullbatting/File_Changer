const path = require('path');
const { pathToFileURL } = require('url');

const appUrl = pathToFileURL(path.join(__dirname, 'index.html')).href;

function webPreferences() {
  return {
    preload: path.join(__dirname, 'preload.js'),
    nodeIntegration: false,
    contextIsolation: true,
    // 기존 파일 변환 모듈과 작업자 프로세스를 격리된 preload에서 사용한다.
    sandbox: false,
    spellcheck: false,
  };
}

function protectWindow(win) {
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-navigate', (event) => event.preventDefault());
  win.webContents.on('will-frame-navigate', (event) => event.preventDefault());
  win.webContents.on('will-attach-webview', (event) => event.preventDefault());
}

function assertTrustedSender(event, win) {
  if (!win || event.sender !== win.webContents ||
      event.senderFrame !== win.webContents.mainFrame || event.senderFrame.url !== appUrl) {
    throw new Error('Untrusted IPC sender');
  }
}

module.exports = { webPreferences, protectWindow, assertTrustedSender };
