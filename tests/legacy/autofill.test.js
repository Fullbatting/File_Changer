// 서식 자동 기입 시나리오: 서식 폴더 등록(파일 이름으로 분류) → 사전에 없는 필드 검토 → 담당자 설정 →
// 인적사항 입력 → 주소로 기관 추천 → 세대주/세대원 → 출력(가짜 한글 인쇄) → 인적사항 지움·임시 파일 지움 →
// (프로덕션과 같은 격리 세계에서 실행: 테스트 코드는 executeJavaScriptInIsolatedWorld, 가짜 인쇄는 autofill-preload.js)
// 디스크에 인적사항이 남지 않았는지 → 한글 파일로 저장 → 백업 내보내기·가져오기. 값은 모두 가짜 예시다.
const { app, BrowserWindow, ipcMain } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');

const REPO = path.resolve(__dirname, '..', '..');
const FIX = process.env.LEGACY_FIX;
const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'autofill-userdata-'));
const saveDir = fs.mkdtempSync(path.join(os.tmpdir(), 'autofill-save-'));
const printed = fs.mkdtempSync(path.join(os.tmpdir(), 'autofill-printed-'));
const backup = path.join(saveDir, 'backup.zip');

const queues = { open: [], save: [], folder: [] };
ipcMain.handle('dialog:openFile', async () => queues.open.shift() || []);
ipcMain.handle('dialog:saveFile', async () => queues.save.shift() || null);
ipcMain.handle('dialog:selectFolder', async () => queues.folder.shift() || null);
ipcMain.handle('app:getUserDataPath', () => userData);
process.on('unhandledRejection', (e) => {
  console.log('UNHANDLED REJECTION: ' + (e && e.stack ? e.stack : e));
  app.exit(2);
});
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const results = [];
const record = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'PASS' : 'FAIL'} - ${name}${ok ? '' : ` :: ${detail}`}`);
};
const sectionOf = (file) => new PizZip(fs.readFileSync(file)).file('Contents/section0.xml').asText();
const allFilesUnder = (dir) => fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => (e.isDirectory() ? allFilesUnder(path.join(dir, e.name)) : [path.join(dir, e.name)]));

app.whenReady().then(async () => {
  // 프로덕션과 같은 보안 설정(격리 세계) + 테스트 preload(가짜 인쇄)
  const win = new BrowserWindow({ show: false, webPreferences: { ...require(path.join(REPO, 'app-security.js')).webPreferences(), preload: path.join(__dirname, 'autofill-preload.js') } });
  await win.loadFile(path.join(REPO, 'index.html'));
  const js = (code) => win.webContents.executeJavaScriptInIsolatedWorld(999, [{ code }]);
  await js(`window.alert = (m) => { window.__alerts = (window.__alerts || []).concat(String(m)); }; window.confirm = () => true; true;`);
  await sleep(600);
  const visible = (id) => js(`!document.getElementById(${JSON.stringify(id)}).hidden`);
  const text = (sel) => js(`(document.querySelector(${JSON.stringify(sel)}) || {}).textContent || ''`);
  const type = (id, value) => js(`(() => { const el = document.getElementById(${JSON.stringify(id)}); el.value = ${JSON.stringify(value)}; el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);

  await js(`document.querySelector('.tab-btn[data-tab="tabAutofill"]').click(); true;`);

  /* ---------- 1. 처음에는 서식 관리 화면 ---------- */
  record('서식이 없으면 서식 관리 화면부터', await visible('screenForms'), 'print screen shown');

  /* ---------- 2. 폴더째 등록: 파일 이름으로 분류 ---------- */
  queues.folder.push(path.join(FIX, 'forms'));
  await js(`document.getElementById('btnAddFolder').click(); true;`);
  await sleep(500);
  const rows = await js(`[...document.querySelectorAll('#importTableBody tr')].map((tr) => ({ file: tr.cells[0].textContent, agency: tr.cells[1].querySelector('input').value, kind: tr.cells[2].querySelector('select').value, applicant: tr.cells[3].querySelector('select').value, title: tr.cells[5].querySelector('input').value, info: tr.cells[6].textContent }))`);
  const byFile = (s) => rows.find((r) => r.file.includes(s)) || {};
  record('등록 확인 창: 서식 3개', (await visible('importModal')) && rows.length === 3, JSON.stringify(rows));
  record('분류: 양주시청 · 긴급복지(국가형) · 세대주/세대원', byFile('(세대주)').agency === '양주시청' && byFile('(세대주)').kind === '긴급복지(국가형)' && byFile('(세대주)').applicant === '세대주' && byFile('(세대원)').applicant === '세대원', JSON.stringify(rows));
  record('분류: 공통 암환자 서식', byFile('암환자').agency === '공통 서식' && byFile('암환자').kind === '암환자 의료비' && byFile('암환자').title === '암환자의료비등록신청서', JSON.stringify(byFile('암환자')));
  record('필드·서류 목차와 사전에 없는 필드 안내', byFile('(세대주)').info.startsWith('필드 23개, 서류 3종') && byFile('암환자').info.includes('사전에 없는 필드: 보호자'), `${byFile('(세대주)').info} / ${byFile('암환자').info}`);
  await js(`document.getElementById('importConfirm').click(); true;`);
  await sleep(500);
  const table = await js(`[...document.querySelectorAll('#formsTableBody tr')].map((tr) => tr.cells[0].textContent + ' | ' + tr.cells[5].textContent)`);
  record('등록 후 목록 3개 · 암환자 서식은 확인 필요', table.length === 3 && table.some((t) => t.includes('암환자') && t.includes('확인 필요')) && table.filter((t) => t.includes('사용 중')).length === 2, JSON.stringify(table));

  /* ---------- 3. 검토: 사전에 없는 필드는 출력 때 입력 ---------- */
  await js(`[...document.querySelectorAll('#formsTableBody tr')].find((tr) => tr.textContent.includes('암환자')).querySelector('button').click(); true;`);
  await sleep(200);
  const reviewRows = await js(`[...document.querySelectorAll('#reviewFields tr')].map((tr) => tr.cells[0].textContent + '=' + tr.querySelector('select').value)`);
  record('검토 창: 사전 필드는 자동, 보호자는 출력할 때 입력', reviewRows.includes('{{이름}}=auto') && reviewRows.some((r) => r.startsWith('{{보호자}}') && r.endsWith('=ask')), JSON.stringify(reviewRows));
  await js(`document.getElementById('reviewSave').click(); true;`);
  await sleep(200);
  const table2 = await js(`[...document.querySelectorAll('#formsTableBody tr')].map((tr) => tr.cells[5].textContent)`);
  record('검토를 저장하면 사용 중', table2.every((t) => t === '사용 중'), JSON.stringify(table2));

  /* ---------- 4. 담당자 설정 ---------- */
  await js(`document.querySelector('[data-screen="screenSettings"]').click(); true;`);
  await type('setStaffName', '[담당자]');
  await type('setStaffPhone', '0000');
  await js(`document.getElementById('btnSaveSettings').click(); true;`);
  await sleep(200);
  const settings = JSON.parse(fs.readFileSync(path.join(userData, 'settings.json'), 'utf8'));
  record('설정 저장(담당자 · 번호)', settings.staffName === '[담당자]' && settings.staffPhone === '0000', JSON.stringify(settings));

  /* ---------- 5. 서류 출력: 주소로 기관 추천, 세대주/세대원 ---------- */
  await js(`document.querySelector('[data-screen="screenPrint"]').click(); true;`);
  await type('in-name', '홍길동');
  await type('in-rrn', '900101-1234567');
  await type('in-addr', '경기도 양주시 [예시로] 00');
  await type('in-phone', '010-0000-0000');
  await sleep(200);
  const chips = await js(`[...document.querySelectorAll('#agencyChips .fa-chip')].map((b) => b.textContent + '|' + b.getAttribute('aria-pressed'))`);
  record('주소(양주시)로 양주시청을 추천해 고름', chips.some((c) => c.startsWith('양주시청') && c.includes('주소로 추천') && c.endsWith('|true')), JSON.stringify(chips));
  record('주민등록번호 안내: 생년월일 6자리 · 성별', (await text('#in-rrn-hint')).startsWith('생년월일 900101, 남'), await text('#in-rrn-hint'));
  let forms = await js(`[...document.querySelectorAll('#formList .fa-form')].map((li) => li.textContent)`);
  record('세대주: 양주시청 세대주 서식 하나만(미리 골라 둠)', forms.length === 1 && forms[0].includes('(세대주)'), JSON.stringify(forms));
  record('고른 서식에 진료기간이 있으면 진료 정보 칸이 보임', (await visible('careBlock')) && (await js(`!document.querySelector('[data-key="careStart"]').hidden && document.querySelector('[data-key="disease"]').hidden`)), 'care fields');
  record('빈칸 안내: 진료 시작일', forms[0].includes('빈칸: 진료 시작일'), forms[0]);
  await type('in-careStart', '2026-01-02');
  await sleep(100);
  forms = await js(`[...document.querySelectorAll('#formList .fa-form')].map((li) => li.textContent)`);
  record('진료 시작일을 넣으면 모두 채워짐', forms[0].includes('모두 채워짐'), forms[0]);
  await js(`document.querySelector('input[name="applicant"][value="세대원"]').click(); true;`);
  await sleep(100);
  forms = await js(`[...document.querySelectorAll('#formList .fa-form')].map((li) => li.textContent)`);
  record('세대원으로 바꾸면 세대원 서식', forms.length === 1 && forms[0].includes('(세대원)'), JSON.stringify(forms));
  await js(`document.querySelector('input[name="applicant"][value="세대주"]').click(); true;`);
  await sleep(100);

  /* ---------- 6. 출력(가짜 한글 인쇄): 채운 파일 확인, 임시 파일·인적사항 지움 ---------- */
  await js(`window.__autofillFakePrint(${JSON.stringify(printed)}); document.getElementById('btnMore').click(); true;`);
  await js(`document.getElementById('btnPrint').click(); true;`);
  await sleep(800);
  const printedFiles = await js('window.__printedFiles');
  record('출력: 서식 1건을 2부로 한글에 보냄', printedFiles.length === 1 && printedFiles[0].copies === 2, JSON.stringify(printedFiles));
  const today = new Date();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const xml = printedFiles.length ? sectionOf(printedFiles[0].copy) : '';
  const cleanText = xml.replace(/<hp:ctrl>.*?<\/hp:ctrl>/g, '').replace(/<[^>]+>/g, '');
  record('채운 서식: 이름·생년월일 6자리·주소·연락처', cleanText.includes('성 명 홍길동 생년월일 900101') && cleanText.includes('(연락처:010-0000-0000)'), cleanText.slice(0, 300));
  record('채운 서식: 작성일 두 자리씩 · 연도는 작성 연도', cleanText.includes(`${today.getFullYear()}년 ${mm[0]}${mm[1]}월 ${dd[0]}${dd[1]}일`), cleanText.slice(0, 400));
  record('채운 서식: 주민등록번호 한 자리씩(서식 칸 순서 4·5·6·1·2·3·7…13) · 진료기간', cleanText.includes('1019001234567') && cleanText.includes('진료기간 26.01.02~') && !cleanText.includes('{{'), cleanText);
  record('채운 서식: 메일머지 필드 표시(fieldBegin/End)는 그대로', (xml.match(/fieldBegin/g) || []).length === 24, String((xml.match(/fieldBegin/g) || []).length));
  record('인쇄에 쓴 임시 파일은 지움', printedFiles.length === 1 && !fs.existsSync(printedFiles[0].temp), printedFiles.length ? printedFiles[0].temp : '');
  record('출력이 끝나면 인적사항을 지움', (await js(`['in-name','in-rrn','in-addr','in-phone','in-careStart'].every((id) => document.getElementById(id).value === '')`)) && (await text('#printStatus')).includes('인적사항을 지웠습니다'), await text('#printStatus'));
  await js(`document.getElementById('printClose').click(); true;`);

  /* ---------- 7. 디스크에 인적사항이 남지 않음 ---------- */
  const leaked = allFilesUnder(userData).filter((f) => {
    const buf = fs.readFileSync(f);
    const hay = buf.toString('utf8') + buf.toString('latin1');
    return hay.includes('홍길동') || hay.includes('900101-1234567') || hay.includes('1234567');
  });
  record('사용자 데이터 폴더(서식·설정)에 인적사항 없음', leaked.length === 0, leaked.join(', '));
  const tempDir = path.join(os.tmpdir(), 'form-autofill');
  const leftovers = fs.existsSync(tempDir) ? fs.readdirSync(tempDir).filter((f) => f.startsWith('form-autofill-')) : [];
  record('임시 폴더에 채운 파일 없음', leftovers.length === 0, leftovers.join(', '));

  /* ---------- 8. 공통 서식(암환자): 담당자 값 · 출력 때 입력 칸 · 한글 파일로 저장 ---------- */
  await type('in-name', '홍길동');
  await type('in-rrn', '900101-1234567');
  await sleep(100);
  await js(`[...document.querySelectorAll('#agencyChips .fa-chip')].find((b) => b.textContent.startsWith('공통 서식')).click(); true;`);
  await sleep(100);
  record('공통 서식의 암환자 서식 하나', (await js(`document.querySelectorAll('#formList .fa-form').length`)) === 1, '');
  record('암환자 서식: 진료 정보(주질병·질병분류기호·진단일) 칸, 출력 때 입력 칸 {{보호자}}', (await js(`!document.querySelector('[data-key="disease"]').hidden && document.querySelector('[data-key="careStart"]').hidden`)) && (await text('#askFields')).includes('{{보호자}}'), await text('#askFields'));
  await type('in-disease', '[예시 상병명]');
  await type('in-diseaseCode', 'C00.0');
  await type('in-diagnosisDate', '2026-01-15');
  await js(`(() => { const el = document.querySelector('#askFields input'); el.value = '[예시 보호자]'; el.dispatchEvent(new Event('input', { bubbles: true })); return true; })()`);
  queues.folder.push(saveDir);
  await js(`document.getElementById('btnSaveFilled').click(); true;`);
  await sleep(700);
  const saved = fs.readdirSync(saveDir).filter((f) => f.endsWith('.hwpx'));
  const savedText = saved.length ? sectionOf(path.join(saveDir, saved[0])).replace(/<hp:ctrl>.*?<\/hp:ctrl>/g, '').replace(/<[^>]+>/g, '') : '';
  record('한글 파일로 저장: 기관_서식 이름으로', saved.length === 1 && saved[0] === '공통 서식_암환자의료비등록신청서.hwpx', saved.join(', '));
  record('저장한 파일: 상병·코드·진단일·담당자 설정·출력 때 입력', savedText.includes('주질병 [예시 상병명] 질병분류기호 C00.0 진단일 26.01.15') && savedText.includes('담당자 [담당자] (0000)') && savedText.includes('보호자 [예시 보호자]'), savedText);
  record('저장한 뒤에는 인적사항을 남겨 둠(이어서 출력할 수 있게)', (await js(`document.getElementById('in-name').value`)) === '홍길동', '');
  await js(`document.getElementById('printClose').click(); document.getElementById('btnClearPerson').click(); true;`);
  record('[새로 입력]으로 모두 지움', (await js(`document.getElementById('in-name').value + document.getElementById('in-disease').value`)) === '', '');

  /* ---------- 9. 백업 내보내기 → 지우기 → 가져오기 ---------- */
  await js(`document.querySelector('[data-screen="screenForms"]').click(); true;`);
  queues.save.push(backup);
  await js(`document.getElementById('btnExport').click(); true;`);
  await sleep(400);
  await js(`[...document.querySelectorAll('#formsTableBody tr')].find((tr) => tr.textContent.includes('(세대원)') || tr.cells[2].textContent === '세대원').querySelector('.danger').click(); true;`);
  await sleep(200);
  const afterDelete = await js(`document.querySelectorAll('#formsTableBody tr').length`);
  queues.open.push([backup]);
  await js(`document.getElementById('btnImport').click(); true;`);
  await sleep(400);
  const afterImport = await js(`document.querySelectorAll('#formsTableBody tr').length`);
  record('백업 내보내기 → 하나 지움 → 가져오면 다시 3개', fs.existsSync(backup) && afterDelete === 2 && afterImport === 3, `${afterDelete} → ${afterImport}`);

  /* ---------- 10. 문서 양식 변환기 탭과 함께 ---------- */
  const tabActive = (id) => js(`document.getElementById(${JSON.stringify(id)}).classList.contains('active')`);
  await js(`document.querySelector('.tab-btn[data-tab="tab1"]').click(); true;`);
  record('상단 탭으로 문서 양식 변환기로 바꿀 수 있음', (await tabActive('tab1')) && !(await tabActive('tabAutofill')) && !(await js(`!!document.getElementById('tab2')`)), '');
  const pageGlobals = await win.webContents.executeJavaScript('({ require: typeof require, process: typeof process })');
  record('웹 페이지 쪽에는 Node가 노출되지 않음(서식 자동 기입도 격리 세계에서 동작)', pageGlobals.require === 'undefined' && pageGlobals.process === 'undefined', JSON.stringify(pageGlobals));

  const passed = results.filter(Boolean).length;
  console.log(`=== ${passed}/${results.length} passed ===`);
  fs.rmSync(userData, { recursive: true, force: true });
  fs.rmSync(saveDir, { recursive: true, force: true });
  fs.rmSync(printed, { recursive: true, force: true });
  app.exit(passed === results.length ? 0 : 1);
});
