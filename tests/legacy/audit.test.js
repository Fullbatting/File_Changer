// 메일머지 자동기입 전체 기능 점검. 실행: electron audit.test.js <phase 1|2|3> (tests/legacy/run.js가 차례로 실행)
// Phase 1: every mail-merge feature. Phase 2: relaunch -> settings persisted. Phase 3: corrupted settings -> recovery.
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

const REPO = path.resolve(__dirname, '..', '..');
const FIX = process.env.LEGACY_FIX;
const PHASE = process.argv[process.argv.length - 1];
const USERDATA = path.join(FIX, 'userdata');
const OUT = path.join(FIX, 'out');
const LOG = path.join(FIX, `audit-phase${PHASE}.log`);
fs.mkdirSync(USERDATA, { recursive: true });
fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(LOG, '');
const flog = (m) => {
  fs.appendFileSync(LOG, m + '\n');
  console.log(m);
};
process.on('unhandledRejection', (e) => {
  // 시나리오 중 예외: 기록하고 바로 끝낸다(멈춘 채 시간 초과까지 기다리지 않도록)
  flog('UNHANDLED REJECTION: ' + (e && e.stack ? e.stack : e));
  app.exit(2);
});

const q = { open: [], save: [], folder: [] };
ipcMain.handle('dialog:openFile', async () => q.open.shift() || []);
ipcMain.handle('dialog:saveFile', async () => q.save.shift() || null);
ipcMain.handle('app:getUserDataPath', () => USERDATA);
ipcMain.handle('dialog:selectFolder', async () => q.folder.shift() || null);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const XLSX = require(path.join(REPO, 'node_modules', 'xlsx'));
const PizZip = require(path.join(REPO, 'node_modules', 'pizzip'));
const fx = (f) => path.join(FIX, f);
const results = [];
const check = (name, ok, detail = '') => {
  results.push(ok);
  flog(`${ok ? 'PASS' : 'FAIL'} - ${name}${ok ? '' : ` :: ${typeof detail === 'string' ? detail : JSON.stringify(detail)}`}`);
};

app.whenReady().then(async () => {
  const win = new BrowserWindow({ show: false, webPreferences: { nodeIntegration: true, contextIsolation: false } });
  const consoleErrors = [];
  win.webContents.on('console-message', (e, lvl, msg) => {
    const level = typeof e.level === 'string' ? { debug: 0, info: 1, warning: 2, error: 3 }[e.level] : lvl;
    const message = e.message !== undefined ? e.message : msg;
    if (level >= 3) consoleErrors.push(message);
  });
  await win.loadFile(path.join(REPO, 'index.html'));
  const js = (code) => win.webContents.executeJavaScript(`{ ${code} }`);
  await js(`window.__alerts = []; window.alert = (m) => { window.__alerts.push(String(m)); }; window.__confirmAnswer = true; window.confirm = (m) => { window.__lastConfirm = String(m); return window.__confirmAnswer; }; true;`);
  await sleep(600);
  const logText = () => js(`document.getElementById('logArea').value`);
  const click = async (id, wait = 250) => { await js(`document.getElementById(${JSON.stringify(id)}).click(); true;`); await sleep(wait); };
  const pick = async (dropId, file, wait = 300) => { q.open.push([file]); await click(dropId, wait); };
  const lastAlert = () => js(`window.__alerts[window.__alerts.length - 1] || ''`);
  const mappingOf = (tbody, attr) => js(`Object.fromEntries(Array.from(document.querySelectorAll('#${tbody} select')).map(s => [s.dataset.${attr}, s.value]))`);
  const readOut = (file) => {
    const wb = XLSX.readFile(file, { cellNF: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref']);
    const cells = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) row.push(ws[XLSX.utils.encode_cell({ r, c })] || null);
      cells.push(row);
    }
    return cells;
  };

  if (PHASE === '1') {
    check('앱 시작 로그', (await logText()).includes('프로그램이 준비되었습니다'));

    // 로그 지우기
    await click('btnClearLog', 100);
    check('로그 지우기', (await logText()) === '');

    // 동의어 사전: 추가 저장 (2단계에서 유지 확인)
    await click('btnManageSynonymsTab2', 150);
    await js(`const t = document.getElementById('synonymGroupsTextarea'); t.value = t.value + '\\n급여, 월급, 봉급'; document.getElementById('synonymManagerSave').click(); true;`);
    await sleep(200);
    const syn = await js(`require(${JSON.stringify(path.join(REPO, 'common.js'))}).autoMatchAll(['월급'], ['급여', '이름'])['월급']`);
    check('동의어 사전 추가가 즉시 자동 매핑에 반영', syn.match === '급여' && syn.reason === 'synonym', syn);

    // 직접 입력: .docx (머리글 변수 + 반복 태그 제외)
    await pick('dropTemplate', fx('letter.docx'));
    await click('btnExtractVars');
    const vars = await js(`Array.from(document.querySelectorAll('#varFormGrid input')).map(i => i.dataset.varName)`);
    check('docx 변수 추출: 본문 + 머리글, 반복 태그 제외', JSON.stringify(vars) === '["이름","가입일","주소","할인율","회사명"]', vars);
    await js(`const v = { 이름: '홍길동', 가입일: '2026-01-15', 주소: '서울\\n중구 & <본관>', 할인율: '10%', 회사명: '한컴' }; document.querySelectorAll('#varFormGrid input').forEach(i => { i.value = v[i.dataset.varName]; }); true;`);
    const manualDocx = path.join(OUT, 'manual.docx');
    q.save.push(manualDocx);
    await click('btnGenerateDocx', 400);
    if (fs.existsSync(manualDocx)) {
      const z = new PizZip(fs.readFileSync(manualDocx));
      const body = z.file('word/document.xml').asText();
      const head = z.file('word/header1.xml').asText();
      const bodyText = body.replace(/<[^>]+>/g, '');
      check('docx 직접입력: 런 분할 변수 치환 + 특수문자', bodyText.includes('홍길동님 가입일 2026-01-15') && bodyText.includes('&amp; &lt;본관&gt;') && body.includes('<w:t xml:space="preserve">홍길동</w:t>'), body.slice(0, 400));
      check('docx 머리글 변수 치환', head.includes('한컴 귀중'), head);
    } else check('docx 직접입력 생성', false, await lastAlert());

    // 직접 입력: .dotx -> 결과는 일반 문서 형식
    await pick('dropTemplate', fx('letter.dotx'));
    await click('btnExtractVars');
    const manualDotx = path.join(OUT, 'from_dotx.docx');
    q.save.push(manualDotx);
    await click('btnGenerateDocx', 400);
    if (fs.existsSync(manualDotx)) {
      const ct = new PizZip(fs.readFileSync(manualDotx)).file('[Content_Types].xml').asText();
      check('.dotx로 만든 결과는 .docx 본문 형식', ct.includes('wordprocessingml.document.main+xml') && !ct.includes('template.main+xml'), ct);
      const head = new PizZip(fs.readFileSync(manualDotx)).file('word/header1.xml').asText();
      check('비워 둔 변수는 "undefined" 대신 빈칸', !head.includes('undefined'), head);
    } else check('.dotx 직접입력 생성', false, await lastAlert());

    // 직접 입력: .hwpx (줄바꿈 -> hp:lineBreak, 특수문자 이스케이프)
    await pick('dropTemplate', fx('letter.hwpx'));
    await click('btnExtractVars');
    await js(`const v = { 이름: '홍길동', 주소: '서울\\n중구 & <본관>' }; document.querySelectorAll('#varFormGrid input').forEach(i => { i.value = v[i.dataset.varName]; }); true;`);
    const manualHwpx = path.join(OUT, 'manual.hwpx');
    q.save.push(manualHwpx);
    await click('btnGenerateDocx', 400);
    if (fs.existsSync(manualHwpx)) {
      const sec = new PizZip(fs.readFileSync(manualHwpx)).file('Contents/section0.xml').asText();
      check('hwpx 직접입력: 특수문자 이스케이프, 런 서식 유지', sec.includes('&amp; &lt;본관&gt;') && sec.includes('<hp:t>홍길동</hp:t></hp:run><hp:run charPrIDRef="1">'), sec);
    } else check('hwpx 직접입력 생성', false, await lastAlert());

    // 일괄 생성: docx 템플릿 + 250행 데이터
    await pick('dropTemplate', fx('letter.docx'));
    await click('btnExtractVars');
    await js(`document.querySelector('input[name="mmMode"][value="batch"]').click(); true;`);
    await pick('dropMailMergeData', fx('mmdata.xlsx'));
    await click('btnLoadMailMergeData', 400);
    const mm = await mappingOf('mailMergeMappingTableBody', 'templateVar');
    check('일괄: 변수↔컬럼 자동 매핑 (회사명은 없음)', JSON.stringify(mm) === JSON.stringify({ 이름: '이름', 가입일: '가입일', 주소: '주소', 할인율: '할인율', 회사명: '' }), mm);
    const mmPrev = await js(`Object.fromEntries(Array.from(document.querySelectorAll('#mailMergeMappingTableBody tr')).map(tr => [tr.querySelector('select').dataset.templateVar, tr.querySelector('.mapping-preview-cell').textContent]))`);
    check('일괄 미리보기: 날짜 YYYY-MM-DD, 비율은 엑셀 표시대로 10%', mmPrev['가입일'] === '2026-01-15 / 2026-01-16' && mmPrev['할인율'] === '10% / 10%', mmPrev);

    const batchDir = path.join(OUT, 'batch');
    fs.rmSync(batchDir, { recursive: true, force: true });
    fs.mkdirSync(path.join(batchDir, '회원100.docx'), { recursive: true }); // a directory in the way -> that row must fail alone

    // 덮어쓰기 확인에서 "아니오" -> 아무것도 안 만듦
    await js(`window.__confirmAnswer = false; true;`);
    q.folder.push(batchDir);
    await click('btnGenerateBatch', 400);
    await click('mailMergeReviewProceed', 600);
    check('기존 파일 덮어쓰기 확인에서 취소하면 생성 안 함', fs.readdirSync(batchDir).length === 1 && (await js(`window.__lastConfirm || ''`)).includes('이름이 같은 파일이 1개'), fs.readdirSync(batchDir));

    // 확인 모달 취소
    await js(`window.__confirmAnswer = true; true;`);
    await click('btnGenerateBatch', 400);
    await click('mailMergeReviewCancel', 200);
    check('일괄 확인 모달 취소', fs.readdirSync(batchDir).length === 1);

    // 실제 생성
    const t0 = Date.now();
    q.folder.push(batchDir);
    await click('btnGenerateBatch', 400);
    const reviewText = await js(`document.getElementById('mailMergeReviewSummary').textContent`);
    check('일괄 확인 모달: 문서 수 + 빈 변수 요약', reviewText.includes('문서 250개') && reviewText.includes('"회사명" 250건'), reviewText);
    await click('mailMergeReviewProceed', 100);
    for (let i = 0; i < 100 && !(await logText()).includes('일괄 생성 완료'); i++) await sleep(200);
    const elapsed = Date.now() - t0;
    const files = fs.readdirSync(batchDir).filter((f) => f.endsWith('.docx'));
    const finalLog = await logText();
    check('일괄 생성: 249건 성공 + 막힌 1건만 실패(오류 격리)', finalLog.includes('총 250건 중 249건 성공, 1건 실패') && finalLog.includes('101행 실패'), finalLog.split('\n').slice(-8).join(' | '));
    check('일괄 생성: 진행 로그', finalLog.includes('진행 중... (100/250)') && finalLog.includes('진행 중... (200/250)'));
    const fileSet = new Set(files);
    check('파일명: 대소문자 중복 구분 Kim/kim(2), 예약어 CON→_CON, 끝의 점 제거 후 중복 (2)', ['Kim.docx', 'kim(2).docx', '_CON.docx', '홍길동.docx', '홍길동(2).docx', '회원5.docx'].every((f) => fileSet.has(f)), files.slice(0, 8));
    const sample = new PizZip(fs.readFileSync(path.join(batchDir, 'Kim.docx')));
    const sb = sample.file('word/document.xml').asText();
    const sh = sample.file('word/header1.xml').asText();
    check('일괄 결과 내용: 날짜 2026-01-15, 할인율 10%, 주소 줄바꿈+특수문자', sb.includes('2026-01-15') && sb.includes('10%') && sb.includes('<w:br/>') && sb.includes('&amp; &lt;본관&gt;'), sb.slice(0, 500));
    check('일괄 결과: 매핑 안 된 머리글 변수는 빈칸 ("undefined" 아님)', !sh.includes('undefined'), sh);
    flog(`INFO - 250건 일괄 생성 소요: ${elapsed}ms`);

    // 일괄 생성: hwpx 템플릿 (데이터의 셀 내 줄바꿈 -> <hp:lineBreak/>)
    await pick('dropTemplate', fx('letter.hwpx'));
    await click('btnExtractVars');
    const hwpxDir = path.join(OUT, 'batch_hwpx');
    fs.rmSync(hwpxDir, { recursive: true, force: true });
    fs.mkdirSync(hwpxDir, { recursive: true });
    q.folder.push(hwpxDir);
    await click('btnGenerateBatch', 400);
    await click('mailMergeReviewProceed', 100);
    for (let i = 0; i < 100 && fs.readdirSync(hwpxDir).length < 250; i++) await sleep(200);
    await sleep(300);
    const kimHwpx = path.join(hwpxDir, 'Kim.hwpx');
    const hsec = fs.existsSync(kimHwpx) ? new PizZip(fs.readFileSync(kimHwpx)).file('Contents/section0.xml').asText() : '';
    check('hwpx 일괄: 250개 + 셀 내 줄바꿈은 <hp:lineBreak/>', fs.readdirSync(hwpxDir).length === 250 && hsec.includes('서울시 중구<hp:lineBreak/>0번지 &amp; &lt;본관&gt;'), hsec || fs.readdirSync(hwpxDir).length);

    // 동의어 사전 창 다시 열기
    await click('btnManageSynonymsTab2', 150);
    check('동의어 사전 버튼', await js(`!document.getElementById('synonymManagerModal').hidden`));
    await click('synonymManagerCancel', 100);
  }

  if (PHASE === '2') {
    await sleep(300);
    await click('btnManageSynonymsTab2', 150);
    check('재시작 후 동의어 사전 유지', await js(`document.getElementById('synonymGroupsTextarea').value.includes('급여, 월급, 봉급')`));
    // 기본값 복원 버튼
    await click('synonymManagerReset', 100);
    const resetText = await js(`document.getElementById('synonymGroupsTextarea').value`);
    check('동의어 사전 기본값 되돌리기', !resetText.includes('월급') && resetText.includes('이름, 성명, 성함'));
    await click('synonymManagerCancel', 100);
  }

  if (PHASE === '3') {
    await sleep(300);
    const log3 = await logText();
    check('손상된 동의어 파일 → 경고 후 기본 사전으로 계속', log3.includes('동의어 사전을 불러오는 중 오류'), log3);
    const syn = await js(`require(${JSON.stringify(path.join(REPO, 'common.js'))}).autoMatchAll(['성명'], ['이름'])['성명']`);
    check('손상 복구 후에도 기본 동의어 매핑 동작', syn.match === '이름', syn);
  }

  check('렌더러 오류 로그 없음', consoleErrors.length === 0, consoleErrors);
  const failed = results.filter((r) => !r).length;
  flog(`=== phase ${PHASE}: ${results.length - failed}/${results.length} passed ===`);
  app.exit(failed ? 1 : 0);
});
