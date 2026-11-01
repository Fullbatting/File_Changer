// 실제 main.js의 창, IPC 및 preload 경로를 테스트한다. 대화상자만 대체한다.
const { app, BrowserWindow, dialog, shell } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');
const PizZip = require('pizzip');

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

  const template = path.join(work, 'template.hwpx');
  const zip = new PizZip();
  zip.file('Contents/section0.xml', '<hs:sec xmlns:hs="urn:section" xmlns:hp="urn:paragraph"><hp:p><hp:run><hp:t>{{이</hp:t></hp:run><hp:run><hp:t>름}}</hp:t></hp:run></hp:p></hs:sec>');
  fs.writeFileSync(template, zip.generate({ type: 'nodebuffer' }));
  open.push([template]);
  await click('dropTemplate');
  await click('btnExtractVars');
  await js(`document.querySelector('#varFormGrid input').value = '홍길동'; true;`);
  const manual = path.join(work, 'manual.hwpx');
  saves.push(manual);
  await click('btnGenerateDocx');
  record('실제 열기·저장 IPC와 HWPX 직접 입력',
    fs.existsSync(manual) && new PizZip(fs.readFileSync(manual)).file('Contents/section0.xml').asText().includes('홍길동') &&
    calls.some((c) => c.kind === 'open' && c.options.properties.includes('openFile')) &&
    calls.some((c) => c.kind === 'save' && c.options.defaultPath.endsWith('.hwpx')), calls);

  const csv = path.join(work, 'data.csv');
  fs.writeFileSync(csv, '\uFEFF이름\n김영희\n이철수\n');
  const out = path.join(work, 'batch');
  fs.mkdirSync(out);
  await js(`const r = document.querySelector('input[name="mmMode"][value="batch"]'); r.checked = true; r.dispatchEvent(new Event('change', { bubbles: true })); true;`);
  open.push([csv]);
  await click('dropMailMergeData');
  await click('btnLoadMailMergeData');
  await click('btnGenerateBatch');
  open.push([out]);
  await click('mailMergeReviewProceed');
  const generated = fs.readdirSync(out).map((file) => new PizZip(fs.readFileSync(path.join(out, file))).file('Contents/section0.xml').asText());
  record('실제 폴더 IPC와 여러 행 메일머지',
    generated.length === 2 && generated.some((xml) => xml.includes('김영희')) &&
    generated.some((xml) => xml.includes('이철수')) &&
    calls.some((c) => c.kind === 'open' && c.options.properties.includes('openDirectory')), calls);

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
