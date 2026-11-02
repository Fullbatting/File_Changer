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
process.on('unhandledRejection', (e) => {
  // 시나리오 중 예외: 기록하고 바로 끝낸다(멈춘 채 시간 초과까지 기다리지 않도록)
  flog('UNHANDLED REJECTION: ' + (e && e.stack ? e.stack : e));
  app.exit(2);
});

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

  /* ---------- 2. HWPX 표를 데이터로 일괄 생성 ---------- */
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
  record('한글 표 데이터 불러오기 (3행)', mmStatus === '3행 불러옴', mmStatus);
  selectFolderQueue.push(batchOutDir);
  await js("document.getElementById('btnGenerateBatch').click()");
  await sleep(300);
  await js("document.getElementById('mailMergeReviewProceed').click()");
  await sleep(700);
  const files = fs.readdirSync(batchOutDir).sort();
  record('일괄 생성: 3명분 문서', JSON.stringify(files) === JSON.stringify(['김영희.docx', '이철수.docx', '홍길동.docx']), JSON.stringify(files));

  /* ---------- 3. 구형 .hwp 안내 (데이터 파일로 고를 때) ---------- */
  await js(`window.__lastAlert = '';`);
  openFileQueue.push([path.join(FIX, 'legacy.hwp')]);
  await js("document.getElementById('dropMailMergeData').click()");
  await sleep(300);
  const hwpAlert = await js('window.__lastAlert || ""');
  record('.hwp 확장자의 손상된 파일은 올바른 문서가 아니라는 안내', hwpAlert.includes('올바른 한글(.hwp) 문서가 아닙니다'), hwpAlert.replace(/\n/g, ' '));

  const failed = results.filter((r) => !r.ok);
  flog(`\n=== ${results.length - failed.length}/${results.length} passed ===`);
  app.exit(failed.length > 0 ? 1 : 0);
}

app.whenReady().then(run);
