// Electron e2e regression test for the renderer.js -> common.js/tab1.js/tab2.js split.
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
  flog('UNHANDLED REJECTION: ' + (err && err.stack ? err.stack : err));
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
    webPreferences: { nodeIntegration: true, contextIsolation: false },
  });
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
  await win.webContents.executeJavaScript(`
    window.alert = (msg) => { window.__lastAlert = msg; };
    window.confirm = (msg) => { window.__lastConfirm = msg; return true; };
    true;
  `);
  flog('CP1: alert stub done');

  await sleep(500); // let the app boot (userDataPath IPC roundtrip + synonym dict load)
  flog('CP2: sleep done');

  const logText = await win.webContents.executeJavaScript("document.getElementById('logArea').value");
  flog('CP3: logText read done, len=' + logText.length);
  record('앱 로드 및 초기 로그에 "준비" 메시지', logText.includes('프로그램이 준비되었습니다'), logText.slice(-200));

  /* ---------------- TAB 1: analyze -> required -> convert -> review -> save ---------------- */

  openFileQueue.push([path.join(FIX, 'sourceA.xlsx')]);
  await win.webContents.executeJavaScript("document.getElementById('dropA').click()");
  await sleep(200);

  openFileQueue.push([path.join(FIX, 'targetB.xlsx')]);
  await win.webContents.executeJavaScript("document.getElementById('dropB').click()");
  await sleep(200);

  await win.webContents.executeJavaScript("document.getElementById('btnAnalyze').click()");
  await sleep(200);

  const analyzeStatus = await win.webContents.executeJavaScript("document.getElementById('analyzeStatus').textContent");
  record('탭1 구조 분석 완료', analyzeStatus === '분석 완료', analyzeStatus);

  // 자동 매핑 확인: 성명<-이름(동의어), 생일<-생년월일(동의어), 연락처<-전화번호(동의어), 이메일<-이메일주소(fuzzy/동의어)
  const mappingSnapshot = await win.webContents.executeJavaScript(`
    Array.from(document.querySelectorAll('#mappingTableBody select')).map(s => [s.dataset.bcolumn, s.value])
  `);
  const mappingObj = Object.fromEntries(mappingSnapshot);
  const expectedAutoMap = { 성명: '이름', 생일: '생년월일', 연락처: '전화번호', 이메일: '이메일주소' };
  const autoMapOk = Object.entries(expectedAutoMap).every(([b, a]) => mappingObj[b] === a);
  record('탭1 동의어 기반 자동 매핑 (성명/생일/연락처/이메일)', autoMapOk, JSON.stringify(mappingObj));

  // 매핑 테이블의 "미리보기" 칸에 실제 원본(A) 데이터 값이 표시되는지 확인
  const previewCells = await win.webContents.executeJavaScript(`
    Array.from(document.querySelectorAll('#mappingTableBody tr')).map(tr => {
      const b = tr.querySelector('td');
      const preview = tr.querySelector('.mapping-preview-cell');
      return [b ? b.textContent : null, preview ? preview.textContent : null];
    })
  `);
  const previewForSungmyeong = previewCells.find(([b]) => b === '성명');
  record(
    '탭1 매핑 테이블 미리보기: "성명" -> "이름" 컬럼의 실제 값 표시',
    !!previewForSungmyeong && previewForSungmyeong[1] === '홍길동 / 김영희',
    JSON.stringify(previewForSungmyeong)
  );

  // 헤더 자동 감지: targetB의 1행은 제목행이므로 건너뛰었는지 로그로 확인
  const logAfterAnalyze = await win.webContents.executeJavaScript("document.getElementById('logArea').value");
  // 타겟(B) 1행 "2026년 명단"은 건너뛰고 2행을 헤더로 — ①단계 인식 결과와 로그에 표시된다
  const recogB = await win.webContents.executeJavaScript("Array.from(document.querySelectorAll('#recogB li')).map((li) => li.textContent)");
  record('탭1 제목행 건너뛰고 헤더 자동 감지', logAfterAnalyze.includes('타겟(B) 파일: 헤더 위 1개 행은 건너뜀: 1행 "2026년 명단"') && recogB[0] === '헤더: 2행(자동)', JSON.stringify(recogB));

  // "비고" 컬럼을 필수로 체크 (매핑 안 된 컬럼 -> 변환 후 빈 값 -> 검증에서 "필수 컬럼 비어있음" 경고가 떠야 함)
  await win.webContents.executeJavaScript(`
    (() => {
      const cb = document.querySelector('#mappingTableBody input[type=checkbox][data-bcolumn="비고"]');
      cb.checked = true;
      cb.dispatchEvent(new Event('change'));
    })();
  `);

  const convertedSavePath = path.join(FIX, 'converted_result.xlsx');
  saveFileQueue.push(convertedSavePath);
  await win.webContents.executeJavaScript("document.getElementById('btnConvert').click()");
  await sleep(300);

  const reviewModalVisible = await win.webContents.executeJavaScript("document.getElementById('convertReviewModal').hidden === false");
  record('탭1 변환 결과 확인 모달 표시', reviewModalVisible, '');

  const reviewSummary = await win.webContents.executeJavaScript("document.getElementById('convertReviewSummary').innerHTML");
  record('탭1 리뷰 모달에 필수컬럼 누락 경고 표시 ("비고")', reviewSummary.includes('비고'), reviewSummary);

  await win.webContents.executeJavaScript("document.getElementById('convertReviewProceed').click()");
  // 저장은 임시 파일에 쓰고 다시 열어 확인(별도 프로세스)한 뒤 원래 이름으로 바꾸므로 파일이 생길 때까지 기다린다
  for (let i = 0; i < 150 && !fs.existsSync(convertedSavePath); i++) await sleep(100);
  await sleep(100);

  const convertedExists = fs.existsSync(convertedSavePath);
  record('탭1 변환 파일 저장됨', convertedExists, convertedSavePath);

  if (convertedExists) {
    const XLSX = require(path.join(REPO, 'node_modules', 'xlsx'));
    const wb = XLSX.readFile(convertedSavePath);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    const firstRowOk = rows[0] && rows[0]['성명'] === '홍길동' && rows[0]['이메일'] === 'hong@example.com' && rows[0]['비고'] === '';
    record('탭1 변환 파일 내용 검증 (매핑된 값 + 빈 비고)', !!firstRowOk, JSON.stringify(rows[0]));
  }

  // 프리셋 저장: 프리셋 이름 모달은 showPresetNameModal()의 onclick 핸들러로 뜨므로, 버튼 클릭 후 입력값을 채우고 저장 클릭
  await win.webContents.executeJavaScript("document.getElementById('btnSavePreset').click()");
  await sleep(200);
  await win.webContents.executeJavaScript(`
    document.getElementById('presetNameInput').value = '테스트프리셋';
    document.getElementById('presetModalSave').click();
  `);
  await sleep(200);
  const presetOptionCount = await win.webContents.executeJavaScript("document.getElementById('presetSelect').options.length");
  record('탭1 프리셋 저장 후 드롭다운에 추가됨', presetOptionCount >= 2, String(presetOptionCount));

  /* ---------------- TAB 2: manual mode (docx) ---------------- */

  await win.webContents.executeJavaScript(`document.querySelector('.tab-btn[data-tab="tab2"]').click()`);
  await sleep(100);

  openFileQueue.push([path.join(FIX, 'template.docx')]);
  await win.webContents.executeJavaScript("document.getElementById('dropTemplate').click()");
  await sleep(200);

  await win.webContents.executeJavaScript("document.getElementById('btnExtractVars').click()");
  await sleep(200);

  const extractStatus = await win.webContents.executeJavaScript("document.getElementById('extractStatus').textContent");
  record('탭2 템플릿 변수 추출 (docx)', extractStatus.includes('2개 변수') || extractStatus.includes('변수 발견'), extractStatus);

  await win.webContents.executeJavaScript(`
    document.querySelectorAll('#varFormGrid input').forEach(inp => {
      inp.value = inp.dataset.varName === '이름' ? '테스트유저' : '2026-10-17';
      inp.dispatchEvent(new Event('input'));
    });
  `);

  const manualSavePath = path.join(FIX, 'manual_result.docx');
  saveFileQueue.push(manualSavePath);
  await win.webContents.executeJavaScript("document.getElementById('btnGenerateDocx').click()");
  await sleep(300);

  const manualExists = fs.existsSync(manualSavePath);
  record('탭2 직접입력 모드 문서 생성됨', manualExists, manualSavePath);
  if (manualExists) {
    const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
    const zip = new PizZip(fs.readFileSync(manualSavePath));
    const xml = zip.file('word/document.xml').asText();
    const substituted = xml.includes('테스트유저') && !xml.includes('{{이름}}');
    record('탭2 직접입력 모드 변수 치환 정확성', substituted, xml.slice(0, 300));
  }

  /* ---------------- TAB 2: batch mode ---------------- */

  await win.webContents.executeJavaScript(`
    document.querySelector('input[name="mmMode"][value="batch"]').click();
  `);
  await sleep(100);

  openFileQueue.push([path.join(FIX, 'mmData.xlsx')]);
  await win.webContents.executeJavaScript("document.getElementById('dropMailMergeData').click()");
  await sleep(200);

  await win.webContents.executeJavaScript("document.getElementById('btnLoadMailMergeData').click()");
  await sleep(200);

  const mmStatus = await win.webContents.executeJavaScript("document.getElementById('mailMergeDataStatus').textContent");
  record('탭2 배치 데이터 불러오기', mmStatus.includes('2행'), mmStatus);

  const mmMapping = await win.webContents.executeJavaScript(`
    Array.from(document.querySelectorAll('#mailMergeMappingTableBody select')).map(s => [s.dataset.templateVar, s.value])
  `);
  const mmMappingOk = mmMapping.some(([v, c]) => v === '이름' && c === '이름') && mmMapping.some(([v, c]) => v === '날짜' && c === '날짜');
  record('탭2 배치 변수<->컬럼 자동 매핑 (exact)', mmMappingOk, JSON.stringify(mmMapping));

  const mmPreviewCells = await win.webContents.executeJavaScript(`
    Array.from(document.querySelectorAll('#mailMergeMappingTableBody tr')).map(tr => {
      const varCell = tr.querySelector('td');
      const preview = tr.querySelector('.mapping-preview-cell');
      return [varCell ? varCell.textContent : null, preview ? preview.textContent : null];
    })
  `);
  const datePreview = mmPreviewCells.find(([v]) => v === '{{날짜}}');
  record(
    '탭2 배치 매핑 미리보기: "날짜" 변수에 사람이 읽기 쉬운 날짜 형식 표시',
    !!datePreview && datePreview[1] === '2026-01-15 / 2026-02-20',
    JSON.stringify(datePreview)
  );

  selectFolderQueue.push(batchOutDir);
  await win.webContents.executeJavaScript("document.getElementById('btnGenerateBatch').click()");
  await sleep(300);

  const batchReviewVisible = await win.webContents.executeJavaScript("document.getElementById('mailMergeReviewModal').hidden === false");
  record('탭2 배치 생성 확인 모달 표시', batchReviewVisible, '');

  await win.webContents.executeJavaScript("document.getElementById('mailMergeReviewProceed').click()");
  await sleep(500);

  const batchFiles = fs.readdirSync(batchOutDir).filter((f) => f.endsWith('.docx'));
  record('탭2 배치 생성: 데이터 2행 -> 문서 2개', batchFiles.length === 2, JSON.stringify(batchFiles));

  if (batchFiles.length === 2) {
    const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
    const sample = new PizZip(fs.readFileSync(path.join(batchOutDir, batchFiles[0])));
    const xml = sample.file('word/document.xml').asText();
    const dateFormatOk = xml.includes('2026-01-15') || xml.includes('2026-02-20');
    record('탭2 배치 생성: 날짜 값이 사람이 읽기 쉬운 형식으로 치환됨', dateFormatOk, xml.slice(0, 300));
  }

  /* ---------------- 동의어 사전 모달 ---------------- */

  await win.webContents.executeJavaScript(`document.querySelector('.tab-btn[data-tab="tab1"]').click()`);
  await sleep(100);
  await win.webContents.executeJavaScript("document.getElementById('btnManageSynonymsTab1').click()");
  await sleep(100);
  const synonymModalVisible = await win.webContents.executeJavaScript("document.getElementById('synonymManagerModal').hidden === false");
  record('동의어 사전 관리 모달 열림 (탭1 버튼)', synonymModalVisible, '');
  const synonymTextHasDefault = await win.webContents.executeJavaScript("document.getElementById('synonymGroupsTextarea').value.includes('이름, 성명, 성함')");
  record('동의어 사전 기본값 로드됨', synonymTextHasDefault, '');
  await win.webContents.executeJavaScript("document.getElementById('synonymManagerCancel').click()");

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
