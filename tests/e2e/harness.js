// tests/e2e/harness.js
// Electron 메인 프로세스에서 앱(index.html)을 열고, 파일 대화상자 등 IPC를 가짜로 바꿔
// 렌더러를 자동으로 조작하는 공통 도구. 각 테스트 파일은 Electron 메인 스크립트로 실행한다:
//   Linux:   TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/<파일>.js
//   Windows: npx electron tests/e2e/<파일>.js

const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const REPO = path.join(__dirname, '..', '..');

function createHarness(name) {
  const queues = { open: [], save: [], folder: [] };
  const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), `wb-e2e-userdata-`));
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), `wb-e2e-work-`));
  const results = [];
  const rendererErrors = [];
  let win = null;

  ipcMain.handle('dialog:openFile', async () => queues.open.shift() || []);
  ipcMain.handle('dialog:saveFile', async () => queues.save.shift() || null);
  ipcMain.handle('dialog:selectFolder', async () => queues.folder.shift() || null);
  ipcMain.handle('app:getUserDataPath', () => userDataDir);

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /** 렌더러에서 코드 실행. { } 블록으로 감싸 전역 const 재선언 오류를 막는다 — 마지막 식의 값이 돌아온다 */
  const js = (code) => win.webContents.executeJavaScript(`{ ${code} }`);

  function record(title, ok, detail) {
    results.push({ title, ok: !!ok, detail });
    const line = `${ok ? 'PASS' : 'FAIL'} - ${title}${!ok && detail !== undefined ? ` :: ${typeof detail === 'string' ? detail : JSON.stringify(detail)}` : ''}`;
    console.log(line);
  }

  async function start({ width = 1360, height = 1000 } = {}) {
    await app.whenReady();
    // 초점·requestAnimationFrame이 실제 앱처럼 동작하도록 창을 띄우고(xvfb 안) 백그라운드 절전을 끈다
    win = new BrowserWindow({ width, height, show: true, webPreferences: { nodeIntegration: true, contextIsolation: false, backgroundThrottling: false } });
    // Electron 35부터 console-message는 이벤트 객체(e.level 문자열, e.message)로 받는다. 오류만 실패로 센다.
    win.webContents.on('console-message', (e) => {
      if (e.level === 'error') rendererErrors.push(e.message);
    });
    await win.loadFile(path.join(REPO, 'index.html'));
    win.focus();
    // 네이티브 alert/confirm은 자동화 중 창을 멈추게 하므로 바로 가짜로 바꾼다
    await js(`window.__alerts = []; window.alert = (m) => { window.__alerts.push(String(m)); }; window.confirm = () => true; true;`);
    await sleep(400);
    return win;
  }

  async function click(id, wait = 120) {
    await js(`document.getElementById(${JSON.stringify(id)}).click(); true;`);
    await sleep(wait);
  }

  /** 파일 선택(가짜 대화상자) → 분석까지 */
  async function loadFiles(sourcePath, targetPath) {
    queues.open.push([sourcePath]);
    await click('dropA', 200);
    queues.open.push([targetPath]);
    await click('dropB', 200);
    await click('btnAnalyze', 400);
  }

  /** 앱 확인 창(#appConfirmModal)이 떠 있으면 내용과 함께 확인/취소를 누른다 */
  async function answerConfirm(ok) {
    const info = await js(`({ open: !document.getElementById('appConfirmModal').hidden, title: document.getElementById('appConfirmTitle').textContent, message: document.getElementById('appConfirmMessage').textContent, items: Array.from(document.querySelectorAll('#appConfirmItems li')).map((li) => li.textContent) })`);
    if (info.open) await click(ok ? 'appConfirmOk' : 'appConfirmCancel', 200);
    return info;
  }

  async function finish() {
    record('렌더러 콘솔 오류 없음', rendererErrors.length === 0, rendererErrors.slice(0, 5));
    const failed = results.filter((r) => !r.ok);
    console.log(`\n[${name}] ${results.length - failed.length}/${results.length} 통과`);
    try {
      fs.rmSync(userDataDir, { recursive: true, force: true });
      fs.rmSync(workDir, { recursive: true, force: true });
    } catch (e) {
      /* 임시 폴더 정리는 실패해도 결과에 영향 없음 */
    }
    app.exit(failed.length > 0 ? 1 : 0);
  }

  process.on('unhandledRejection', (err) => {
    console.log(`UNHANDLED: ${err && err.stack ? err.stack : err}`);
    app.exit(2);
  });

  return {
    queues,
    userDataDir,
    workDir,
    start,
    js,
    sleep,
    click,
    loadFiles,
    answerConfirm,
    record,
    finish,
    get win() {
      return win;
    },
  };
}

module.exports = { createHarness, REPO };
