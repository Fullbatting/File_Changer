// 실제 main.js의 창, IPC 및 preload 경로를 테스트한다. 대화상자만 대체한다.
const { app, BrowserWindow, dialog, shell } = require('electron');
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
dialog.showSaveDialog = async (win, options) => {
  calls.push({ kind: 'save', options });
  const filePath = saves.shift();
  return { canceled: !filePath, filePath };
};
shell.showItemInFolder = (filePath) => calls.push({ kind: 'show', filePath });
require('../../main');

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

  const source = path.join(work, 'source.csv');
  const target = path.join(work, 'target.csv');
  fs.writeFileSync(source, '\uFEFF이름,부서\n홍길동,개발\n김영희,영업\n');
  fs.writeFileSync(target, '\uFEFF이름,부서\n');
  open.push([source]);
  await click('dropA');
  open.push([target]);
  await click('dropB');
  await click('btnAnalyze');
  await click('btnSavePreset');
  await js(`document.getElementById('presetNameInput').value = '보안회귀';
    document.getElementById('presetModalSave').click(); true;`);
  await sleep(200);
  record('실제 사용자 데이터 IPC와 프리셋 저장', fs.existsSync(path.join(userData, 'mapping-presets.json')));

  const converted = path.join(work, 'converted.xlsx');
  saves.push(converted);
  await click('btnConvert');
  await click('convertReviewProceed');
  for (let i = 0; i < 100 && !fs.existsSync(converted); i++) await sleep(100);
  await click('btnShowSaved');
  record('실제 변환 저장과 파일 탐색기 IPC', fs.existsSync(converted) &&
    calls.some((c) => c.kind === 'show' && c.filePath === converted), calls);

  // 서식 자동 기입 탭: 서식 파일 추가(열기 IPC) → 등록, 폴더째 추가(폴더 IPC)
  const form = path.join(repo, 'tests', 'fixtures', 'forms', '암환자_의료비_등록_신청서.hwp');
  open.push([form]);
  await click('btnAddFiles');
  await click('importConfirm');
  const formRows = await js(`document.querySelectorAll('#formsTableBody tr[data-id]').length`);
  const emptyDir = path.join(work, 'empty');
  fs.mkdirSync(emptyDir);
  open.push([emptyDir]);
  await click('btnAddFolder');
  record('서식 자동 기입: 실제 열기·폴더 IPC와 사용자 데이터 폴더의 서식 보관함', formRows === 1 &&
    calls.some((c) => c.kind === 'open' && c.options.properties.includes('openDirectory')) &&
    fs.readdirSync(userData).includes('forms'), { formRows, calls });

  const before = calls.length;
  const outsider = new BrowserWindow({
    show: false,
    webPreferences: { preload: path.join(__dirname, 'untrusted-preload.js'), nodeIntegration: false, contextIsolation: true, sandbox: false },
  });
  await outsider.loadFile(path.join(repo, 'index.html'));
  let denied;
  for (let i = 0; i < 50; i++) {
    denied = await outsider.webContents.executeJavaScriptInIsolatedWorld(999, [{ code: 'window.__ipcResults' }]);
    if (denied) break;
    await sleep(20);
  }
  record('같은 앱 URL의 다른 창도 모든 실제 IPC에서 거부',
    denied && denied.length === 5 && denied.every(Boolean) && calls.length === before, denied);
  outsider.destroy();
}

run().catch((err) => record('시나리오 예외 없음', false, err.stack)).finally(() => {
  console.log(`\n[production-security] ${results.filter(Boolean).length}/${results.length} 통과`);
  const failed = results.length === 0 || results.some((ok) => !ok);
  fs.rmSync(work, { recursive: true, force: true });
  app.exit(failed ? 1 : 0);
});
