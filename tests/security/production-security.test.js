// 실제 main.js의 창, IPC 및 preload 경로를 테스트한다. 대화상자만 대체한다.
const { app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');

const repo = path.join(__dirname, '..', '..');
const work = fs.mkdtempSync(path.join(os.tmpdir(), 'production-security-'));
const userData = path.join(work, 'user-data');
fs.mkdirSync(userData);
app.setPath('userData', userData);

const open = [];
const saves = [];
const calls = [];
dialog.showOpenDialog = async (win, options) => {
  calls.push({ kind: 'open', options });
  const paths = open.shift() || [];
  return { canceled: paths.length === 0, filePaths: paths };
};
const errorBoxes = [];
dialog.showErrorBox = (title, text) => errorBoxes.push({ title, text });
dialog.showSaveDialog = async (win, options) => {
  calls.push({ kind: 'save', options });
  const filePath = saves.shift();
  return { canceled: !filePath, filePath };
};
require('../../src/main/main');

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const results = [];
function record(title, ok, detail) {
  results.push(!!ok);
  console.log(`${ok ? 'PASS' : 'FAIL'} - ${title}${!ok ? ` :: ${JSON.stringify(detail)}` : ''}`);
}

async function run() {
  await app.whenReady();
  while (BrowserWindow.getAllWindows().length === 0) await sleep(20);
  const win = BrowserWindow.getAllWindows()[0];
  if (win.webContents.isLoading()) await new Promise((resolve) => win.webContents.once('did-finish-load', resolve));
  const js = (code) => win.webContents.executeJavaScriptInIsolatedWorld(999, [{ code: `{ ${code} }` }]);
  const click = async (id) => { await js(`document.getElementById(${JSON.stringify(id)}).click(); true;`); await sleep(200); };
  await js(`window.alert = () => {}; window.confirm = () => true; true;`);
  await sleep(300);

  const prefs = win.webContents.getLastWebPreferences();
  const globals = await win.webContents.executeJavaScript(`({
    require: typeof require, process: typeof process, ipc: typeof window.ipcRenderer,
    ready: document.getElementById('logArea').value.includes('프로그램이 준비되었습니다'),
    csp: document.querySelector('meta[http-equiv="Content-Security-Policy"]').content
  })`);
  record('실제 앱의 격리 설정과 페이지 Node/IPC 차단',
    prefs.nodeIntegration === false && prefs.contextIsolation === true && globals.ready &&
    globals.require === 'undefined' && globals.process === 'undefined' && globals.ipc === 'undefined' &&
    globals.csp.includes("script-src 'none'"), { prefs, globals });

  await win.webContents.executeJavaScript(`window.__executed = false;
    const script = document.createElement('script'); script.textContent = 'window.__executed = true'; document.body.appendChild(script);
    const image = document.createElement('img'); image.setAttribute('onerror', 'window.__executed = true'); image.src = 'data:image/png;base64,invalid'; document.body.appendChild(image); true;`);
  await sleep(200);
  record('실제 페이지 CSP가 스크립트와 이벤트 속성을 차단', !(await win.webContents.executeJavaScript('window.__executed')));

  // 서식 파일 추가(열기 IPC) → 등록 확인 창 → 등록
  const form = path.join(repo, 'tests', 'fixtures', 'forms', '암환자_의료비_등록_신청서.hwp');
  open.push([form]);
  await click('btnAddFiles');
  await click('importConfirm');
  const rows = await js(`document.querySelectorAll('#formsTableBody tr[data-id]').length`);
  record('실제 열기 IPC로 서식 등록', rows === 1 && calls.some((c) => c.kind === 'open' && c.options.properties.includes('openFile') && c.options.properties.includes('multiSelections')), { rows, calls });

  // 폴더째 추가(폴더 IPC): 한글 서식이 없는 폴더면 안내만
  const empty = path.join(work, 'empty');
  fs.mkdirSync(empty);
  open.push([empty]);
  await click('btnAddFolder');
  record('실제 폴더 IPC', calls.some((c) => c.kind === 'open' && c.options.properties.includes('openDirectory')), calls);

  // 백업 내보내기(저장 IPC)
  const backup = path.join(work, 'backup.zip');
  saves.push(backup);
  await click('btnExport');
  for (let i = 0; i < 50 && !fs.existsSync(backup); i++) await sleep(20);
  record('실제 저장 IPC로 서식 백업', fs.existsSync(backup) && calls.some((c) => c.kind === 'save' && c.options.defaultPath.endsWith('.zip')), calls);

  // 서식 보관함은 사용자 데이터 폴더(IPC로 받은 곳) 안에 둔다
  const stored = fs.readdirSync(userData);
  record('실제 사용자 데이터 IPC와 서식 보관함 위치', stored.includes('forms') || stored.includes('form-library'), stored);

  // 검은 창 없이 실행되므로 문제는 오류 창으로: 화면 준비 실패(메인), 화면 코드에서 처리 못 한 오류(화면)
  win.webContents.emit('preload-error', {}, 'preload.js', new Error('[시험] 모듈을 찾을 수 없음'));
  await js(`window.__alerts = []; window.alert = (m) => window.__alerts.push(String(m)); setTimeout(() => { throw new Error('[시험] 처리 못 한 오류'); }); Promise.reject(new Error('[시험] 처리 못 한 약속')); true;`);
  await sleep(300);
  const alerts = await js('window.__alerts');
  record('문제가 생기면 오류 창으로 알림(검은 창 없이 실행)',
    errorBoxes.some((b) => b.title.includes('오류') && b.text.includes('화면을 준비하지 못했습니다') && b.text.includes('[시험] 모듈')) &&
    alerts.some((a) => a.includes('예상하지 못한 오류') && a.includes('[시험] 처리 못 한 오류')) &&
    alerts.some((a) => a.includes('[시험] 처리 못 한 약속')), { errorBoxes, alerts });

  const before = calls.length;
  const outsider = new BrowserWindow({
    show: false,
    webPreferences: { preload: path.join(__dirname, 'untrusted-preload.js'), nodeIntegration: false, contextIsolation: true, sandbox: false },
  });
  await outsider.loadFile(path.join(repo, 'src', 'renderer', 'index.html'));
  let denied;
  for (let i = 0; i < 50; i++) {
    denied = await outsider.webContents.executeJavaScriptInIsolatedWorld(999, [{ code: 'window.__ipcResults' }]);
    if (denied) break;
    await sleep(20);
  }
  record('같은 앱 URL의 다른 창도 모든 실제 IPC에서 거부',
    denied && denied.length === 4 && denied.every(Boolean) && calls.length === before, denied);
  outsider.destroy();
}

run().catch((err) => record('시나리오 예외 없음', false, err.stack)).finally(() => {
  console.log(`\n[production-security] ${results.filter(Boolean).length}/${results.length} 통과`);
  const failed = results.length === 0 || results.some((ok) => !ok);
  fs.rmSync(work, { recursive: true, force: true });
  app.exit(failed ? 1 : 0);
});
