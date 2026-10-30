// Electron e2e test: HWPX tables as data source for Tab1 (converter) and Tab2 (batch mail merge).
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const REPO = path.resolve(__dirname, '..', '..');
const FIX = process.env.LEGACY_FIX;
const LOG_FILE = path.join(FIX, 'e2e-hwpx-result.log');
fs.writeFileSync(LOG_FILE, '');
const flog = (m) => {
  fs.appendFileSync(LOG_FILE, m + '\n');
  console.log(m);
};
process.on('unhandledRejection', (e) => flog('UNHANDLED REJECTION: ' + (e && e.stack ? e.stack : e)));

const openFileQueue = [];
const saveFileQueue = [];
const selectFolderQueue = [];
const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-userdata-'));
const batchOutDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-batchout-'));

ipcMain.handle('dialog:openFile', async () => openFileQueue.shift() || []);
ipcMain.handle('dialog:saveFile', async () => saveFileQueue.shift() || null);
ipcMain.handle('app:getUserDataPath', () => userDataDir);
ipcMain.handle('dialog:selectFolder', async () => selectFolderQueue.shift() || null);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function run() {
  const win = new BrowserWindow({ show: false, webPreferences: { nodeIntegration: true, contextIsolation: false } });
  await win.loadFile(path.join(REPO, 'index.html'));
  const js = (code) => win.webContents.executeJavaScript(code);

  const results = [];
  const record = (name, ok, detail) => {
    results.push({ name, ok });
    flog(`${ok ? 'PASS' : 'FAIL'} - ${name}${detail ? ' :: ' + detail : ''}`);
  };

  await js(`window.alert = (m) => { window.__lastAlert = m; }; window.confirm = () => true; true;`);
  await sleep(500);

  const staff = path.join(FIX, 'staff.hwpx');

  /* ---------- 1. 표 파싱 자체 (공통 모듈 직접 호출) ---------- */
  const parsed = await js(`(() => {
    const common = require(${JSON.stringify(path.join(REPO, 'common.js'))});
    const XLSX = require(${JSON.stringify(path.join(REPO, 'node_modules', 'xlsx'))});
    const wb = common.readWorkbookSmart(${JSON.stringify(staff)});
    const sheets = {};
    wb.SheetNames.forEach((n) => { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }); });
    const approvalObjects = XLSX.utils.sheet_to_json(wb.Sheets['표 1'], { defval: '' });
    return { names: wb.SheetNames, preferred: wb.preferredSheetName, sheets, approvalObjects };
  })()`);

  record('표 4개 인식 (빈 표 제외, 중첩 표 포함, 문서 순서)', JSON.stringify(parsed.names) === JSON.stringify(['표 1', '표 2', '표 3', '표 4']), JSON.stringify(parsed.names));
  record('행이 가장 많은 데이터 표(표 2)를 기본 선택으로 제안', parsed.preferred === '표 2', parsed.preferred);
  const expectedMain = [
    ['2026년 하반기 직원 명단', '', '', ''],
    ['이름', '연락처', '', '부서'],
    ['이름', '전화번호', '이메일', '부서'],
    ['홍길동', '010-1111-2222', 'hong@example.com', '영업팀'],
    ['김영희', '010-3333-4444', 'kim&co@example.com', '영업팀'],
    ['이철수', '010-5555-6666', 'lee@example.com', '개발\n2팀'],
  ];
  record(
    '데이터 표: 병합셀 배치 + 세로병합 채움 + &amp; 해석 + 셀 내 줄바꿈',
    JSON.stringify(parsed.sheets['표 2']) === JSON.stringify(expectedMain),
    JSON.stringify(parsed.sheets['표 2'])
  );
  record('결재란 표(표 1): 헤더 + 빈 서명칸 행', JSON.stringify(parsed.sheets['표 1']) === JSON.stringify([['담당', '팀장', '부장'], ['', '', '']]), JSON.stringify(parsed.sheets['표 1']));
  record('완전히 빈 행은 데이터 행으로 읽히지 않음 (엑셀과 동일)', parsed.approvalObjects.length === 0, JSON.stringify(parsed.approvalObjects));
  record('중첩 표: 바깥 셀에는 바깥 글자만', JSON.stringify(parsed.sheets['표 3']) === JSON.stringify([['참고', '메모']]), JSON.stringify(parsed.sheets['표 3']));
  record('중첩 표: 안쪽 표는 별도 표로', JSON.stringify(parsed.sheets['표 4']) === JSON.stringify([['내부표']]), JSON.stringify(parsed.sheets['표 4']));

  /* ---------- 2. 탭1: HWPX(A) -> 엑셀 양식(B) 변환 ---------- */
  openFileQueue.push([staff]);
  await js("document.getElementById('dropA').click()");
  await sleep(300);
  const selectorState = await js(`({
    visible: !document.getElementById('sheetRowA').hidden,
    label: document.querySelector('#sheetRowA label').textContent,
    value: document.getElementById('sheetSelectA').value,
  })`);
  record('탭1 표 선택 UI: 보이고, "표 선택" 라벨, 표 2 기본 선택', selectorState.visible && selectorState.label === '표 선택' && selectorState.value === '표 2', JSON.stringify(selectorState));

  openFileQueue.push([path.join(FIX, 'targetB.xlsx')]);
  await js("document.getElementById('dropB').click()");
  await sleep(300);
  const labelB = await js("document.querySelector('#sheetRowB label').textContent");
  record('엑셀 파일(B)은 라벨이 "시트 선택" 그대로', labelB === '시트 선택', labelB);

  await js("document.getElementById('btnAnalyze').click()");
  await sleep(300);
  const logAfterAnalyze = await js("document.getElementById('logArea').value");
  const recogA = await js("Array.from(document.querySelectorAll('#recogA li')).map((li) => li.textContent)");
  record(
    '2단 헤더 표에서 아래 헤더 행(3행) 자동 감지',
    logAfterAnalyze.includes('원본(A) 파일: 헤더 위 2개 행은 데이터에서 뺌: 1행 "2026년 하반기 직원 명단", 2행 "이름, 연락처, 부서"') && recogA[0].startsWith('헤더: 3행(자동)'),
    JSON.stringify(recogA)
  );
  const mapping = Object.fromEntries(await js(`Array.from(document.querySelectorAll('#mappingTableBody select')).map(s => [s.dataset.bcolumn, s.value])`));
  record('탭1 자동 매핑 (성명←이름, 연락처←전화번호, 이메일←이메일)', mapping['성명'] === '이름' && mapping['연락처'] === '전화번호' && mapping['이메일'] === '이메일', JSON.stringify(mapping));
  const preview = await js(`(() => { const tr = Array.from(document.querySelectorAll('#mappingTableBody tr')).find(t => t.querySelector('td').textContent === '성명'); return tr.querySelector('.mapping-preview-cell').textContent; })()`);
  record('탭1 미리보기에 한글 표의 실제 값', preview === '홍길동 / 김영희', preview);

  const outPath = path.join(FIX, 'hwpx_converted.xlsx');
  if (fs.existsSync(outPath)) fs.unlinkSync(outPath);
  saveFileQueue.push(outPath);
  await js("document.getElementById('btnConvert').click()");
  await sleep(300);
  await js("document.getElementById('convertReviewProceed').click()");
  // 저장은 임시 파일에 쓰고 다시 열어 확인(별도 프로세스)한 뒤 원래 이름으로 바꾸므로 파일이 생길 때까지 기다린다
  for (let i = 0; i < 150 && !fs.existsSync(outPath); i++) await sleep(100);
  await sleep(100);
  if (fs.existsSync(outPath)) {
    const XLSX = require(path.join(REPO, 'node_modules', 'xlsx'));
    const wb = XLSX.readFile(outPath);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    const ok =
      rows.length === 3 &&
      rows[0]['성명'] === '홍길동' &&
      rows[1]['이메일'] === 'kim&co@example.com' &&
      rows[2]['연락처'] === '010-5555-6666';
    record('탭1 변환 결과 엑셀: 3행, 값 정확', ok, JSON.stringify(rows));
  } else {
    record('탭1 변환 결과 엑셀 저장됨', false, await js('window.__lastAlert || ""'));
  }

  /* ---------- 3. 탭2: HWPX 표를 데이터로 일괄 생성 ---------- */
  await js(`document.querySelector('.tab-btn[data-tab="tab2"]').click()`);
  openFileQueue.push([path.join(FIX, 'template.docx')]);
  await js("document.getElementById('dropTemplate').click()");
  await sleep(200);
  await js("document.getElementById('btnExtractVars').click()");
  await sleep(200);
  await js(`document.querySelector('input[name="mmMode"][value="batch"]').click()`);
  openFileQueue.push([staff]);
  await js("document.getElementById('dropMailMergeData').click()");
  await sleep(300);
  await js("document.getElementById('btnLoadMailMergeData').click()");
  await sleep(300);
  const mmStatus = await js("document.getElementById('mailMergeDataStatus').textContent");
  record('탭2 한글 표 데이터 불러오기 (3행)', mmStatus === '3행 불러옴', mmStatus);
  selectFolderQueue.push(batchOutDir);
  await js("document.getElementById('btnGenerateBatch').click()");
  await sleep(300);
  await js("document.getElementById('mailMergeReviewProceed').click()");
  await sleep(700);
  const files = fs.readdirSync(batchOutDir).sort();
  record('탭2 일괄 생성: 3명분 문서', JSON.stringify(files) === JSON.stringify(['김영희.docx', '이철수.docx', '홍길동.docx']), JSON.stringify(files));

  /* ---------- 4. 구형 .hwp 안내 ---------- */
  await js(`document.querySelector('.tab-btn[data-tab="tab1"]').click(); window.__lastAlert = '';`);
  openFileQueue.push([path.join(FIX, 'legacy.hwp')]);
  await js("document.getElementById('dropA').click()");
  await sleep(300);
  const hwpAlert = await js('window.__lastAlert || ""');
  record('.hwp 확장자의 손상된 파일은 올바른 문서가 아니라는 안내', hwpAlert.includes('올바른 한글(.hwp) 문서가 아닙니다'), hwpAlert.replace(/\n/g, ' '));

  const failed = results.filter((r) => !r.ok);
  flog(`\n=== ${results.length - failed.length}/${results.length} passed ===`);
  app.exit(failed.length > 0 ? 1 : 0);
}

app.whenReady().then(run);
