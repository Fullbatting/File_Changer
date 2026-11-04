// 메일머지 자동기입 기본 흐름 회귀 테스트: 변수 추출·직접 입력 생성·데이터 파일 일괄 생성·동의어 사전 창
// Run with: xvfb-run -a <repo>/node_modules/.bin/electron --no-sandbox --disable-gpu <this file>
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const REPO = path.resolve(__dirname, '..', '..');
const FIX = process.env.LEGACY_FIX;

let win;
let openFileQueue = []; // array of {paths} to return in order for dialog:openFile
let saveFileQueue = [];
let selectFolderQueue = [];

const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-convert-userdata-'));
const batchOutDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc-convert-batchout-'));

const LOG_FILE = path.join(FIX, 'e2e-result.log');
fs.writeFileSync(LOG_FILE, '');
function flog(msg) {
  fs.appendFileSync(LOG_FILE, msg + '\n');
  console.log(msg);
}
process.on('unhandledRejection', (err) => {
  // 시나리오 중 예외: 기록하고 바로 끝낸다(멈춘 채 시간 초과까지 기다리지 않도록)
  flog('UNHANDLED REJECTION: ' + (err && err.stack ? err.stack : err));
  app.exit(2);
});

ipcMain.handle('dialog:openFile', async () => {
  if (openFileQueue.length === 0) return [];
  return openFileQueue.shift();
});
ipcMain.handle('dialog:saveFile', async () => {
  if (saveFileQueue.length === 0) return null;
  return saveFileQueue.shift();
});
ipcMain.handle('app:getUserDataPath', () => userDataDir);
ipcMain.handle('dialog:selectFolder', async () => {
  if (selectFolderQueue.length === 0) return null;
  return selectFolderQueue.shift();
});

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function run() {
  win = new BrowserWindow({
    show: false,
    webPreferences: require('../../app-security').webPreferences(),
  });
  const js = (code) => win.webContents.executeJavaScriptInIsolatedWorld(999, [{ code }]);
  flog('BrowserWindow created, loading index.html...');
  await win.loadFile(path.join(REPO, 'index.html'));
  flog('index.html loaded');

  const results = [];
  const record = (name, ok, detail) => {
    results.push({ name, ok, detail });
    flog(`${ok ? 'PASS' : 'FAIL'} - ${name}${detail ? ' :: ' + detail : ''}`);
  };
  flog('run() started');

  // Capture renderer console errors
  const rendererErrors = [];
  win.webContents.on('console-message', (e, lvl, msg) => {
    const level = typeof e.level === 'string' ? { debug: 0, info: 1, warning: 2, error: 3 }[e.level] : lvl;
    const message = e.message !== undefined ? e.message : msg;
    if (level >= 3) rendererErrors.push(message);
  });

  // 네이티브 alert()/confirm()은 자동화 환경에서 클릭할 수 없어 렌더러를 영원히
  // 멈추게 하므로, 테스트 중에는 가짜로 교체해 둔다 (실제 앱 동작에는 영향 없음).
  await js(`
    window.alert = (msg) => { window.__lastAlert = msg; };
    window.confirm = (msg) => { window.__lastConfirm = msg; return true; };
    true;
  `);
  flog('CP1: alert stub done');

  await sleep(500); // let the app boot (userDataPath IPC roundtrip + synonym dict load)
  flog('CP2: sleep done');

  const logText = await js("document.getElementById('logArea').value");
  flog('CP3: logText read done, len=' + logText.length);
  record('앱 로드 및 초기 로그에 "준비" 메시지', logText.includes('프로그램이 준비되었습니다'), logText.slice(-200));

  /* ---------------- 직접 입력 (docx) ---------------- */

  openFileQueue.push([path.join(FIX, 'template.docx')]);
  await js("document.getElementById('dropTemplate').click()");
  await sleep(200);

  await js("document.getElementById('btnExtractVars').click()");
  await sleep(200);

  const extractStatus = await js("document.getElementById('extractStatus').textContent");
  record('템플릿 변수 추출 (docx)', extractStatus.includes('2개 변수') || extractStatus.includes('변수 발견'), extractStatus);

  await js(`
    document.querySelectorAll('#varFormGrid input').forEach(inp => {
      inp.value = inp.dataset.varName === '이름' ? '테스트유저' : '2026-10-17';
      inp.dispatchEvent(new Event('input'));
    });
  `);

  const manualSavePath = path.join(FIX, 'manual_result.docx');
  saveFileQueue.push(manualSavePath);
  await js("document.getElementById('btnGenerateDocx').click()");
  await sleep(300);

  const manualExists = fs.existsSync(manualSavePath);
  record('직접입력 모드 문서 생성됨', manualExists, manualSavePath);
  if (manualExists) {
    const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
    const zip = new PizZip(fs.readFileSync(manualSavePath));
    const xml = zip.file('word/document.xml').asText();
    const substituted = xml.includes('테스트유저') && !xml.includes('{{이름}}');
    record('직접입력 모드 변수 치환 정확성', substituted, xml.slice(0, 300));
  }

  /* ---------------- TAB 2: batch mode ---------------- */

  await js(`
    document.querySelector('input[name="mmMode"][value="batch"]').click();
  `);
  await sleep(100);

  openFileQueue.push([path.join(FIX, 'mmData.xlsx')]);
  await js("document.getElementById('dropMailMergeData').click()");
  await sleep(200);

  await js("document.getElementById('btnLoadMailMergeData').click()");
  await sleep(200);

  const mmStatus = await js("document.getElementById('mailMergeDataStatus').textContent");
  record('배치 데이터 불러오기', mmStatus.includes('2행'), mmStatus);

  const mmMapping = await js(`
    Array.from(document.querySelectorAll('#mailMergeMappingTableBody select')).map(s => [s.dataset.templateVar, s.value])
  `);
  const mmMappingOk = mmMapping.some(([v, c]) => v === '이름' && c === '이름') && mmMapping.some(([v, c]) => v === '날짜' && c === '날짜');
  record('배치 변수<->컬럼 자동 매핑 (exact)', mmMappingOk, JSON.stringify(mmMapping));

  const mmPreviewCells = await js(`
    Array.from(document.querySelectorAll('#mailMergeMappingTableBody tr')).map(tr => {
      const varCell = tr.querySelector('td');
      const preview = tr.querySelector('.mapping-preview-cell');
      return [varCell ? varCell.textContent : null, preview ? preview.textContent : null];
    })
  `);
  const datePreview = mmPreviewCells.find(([v]) => v === '{{날짜}}');
  record(
    '배치 매핑 미리보기: "날짜" 변수에 사람이 읽기 쉬운 날짜 형식 표시',
    !!datePreview && datePreview[1] === '2026-01-15 / 2026-02-20',
    JSON.stringify(datePreview)
  );

  selectFolderQueue.push(batchOutDir);
  await js("document.getElementById('btnGenerateBatch').click()");
  await sleep(300);

  const batchReviewVisible = await js("document.getElementById('mailMergeReviewModal').hidden === false");
  record('배치 생성 확인 모달 표시', batchReviewVisible, '');

  await js("document.getElementById('mailMergeReviewProceed').click()");
  await sleep(500);

  const batchFiles = fs.readdirSync(batchOutDir).filter((f) => f.endsWith('.docx'));
  record('배치 생성: 데이터 2행 -> 문서 2개', batchFiles.length === 2, JSON.stringify(batchFiles));

  if (batchFiles.length === 2) {
    const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
    const sample = new PizZip(fs.readFileSync(path.join(batchOutDir, batchFiles[0])));
    const xml = sample.file('word/document.xml').asText();
    const dateFormatOk = xml.includes('2026-01-15') || xml.includes('2026-02-20');
    record('배치 생성: 날짜 값이 사람이 읽기 쉬운 형식으로 치환됨', dateFormatOk, xml.slice(0, 300));
  }

  /* ---------------- 동의어 사전 모달 ---------------- */

  await js("document.getElementById('btnManageSynonymsTab2').click()");
  await sleep(100);
  const synonymModalVisible = await js("document.getElementById('synonymManagerModal').hidden === false");
  record('동의어 사전 관리 모달 열림', synonymModalVisible, '');
  const synonymTextHasDefault = await js("document.getElementById('synonymGroupsTextarea').value.includes('이름, 성명, 성함')");
  record('동의어 사전 기본값 로드됨', synonymTextHasDefault, '');
  await js("document.getElementById('synonymManagerCancel').click()");

  /* ---------------- 콘솔 에러 체크 ---------------- */
  record('렌더러 콘솔 에러 없음', rendererErrors.length === 0, JSON.stringify(rendererErrors));

  const failed = results.filter((r) => !r.ok);
  flog('\n=== SUMMARY ===');
  flog(`${results.length - failed.length}/${results.length} passed`);
  if (failed.length > 0) {
    flog('FAILED: ' + JSON.stringify(failed.map((f) => f.name)));
  }
  flog('run() finished, exiting');

  await sleep(100);
  app.exit(failed.length > 0 ? 1 : 0);
}

app.whenReady().then(run);
