// Full feature audit for the Electron app. Usage: electron audit.js <phase 1|2|3>
// Phase 1: every feature of both tabs. Phase 2: relaunch -> settings persisted. Phase 3: corrupted settings -> recovery.
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
process.on('unhandledRejection', (e) => flog('UNHANDLED REJECTION: ' + (e && e.stack ? e.stack : e)));

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
  const tab = (t) => js(`document.querySelector('.tab-btn[data-tab="${t}"]').click(); true;`);
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
    /* ================= TAB 1 ================= */
    check('앱 시작 로그', (await logText()).includes('프로그램이 준비되었습니다'));

    // 드래그앤드롭 (A)
    // Real drag-and-drop: a genuine disk-backed File (via DevTools DOM.setFileInputFiles), dropped on the zone.
    await js(`const inp = document.createElement('input'); inp.type = 'file'; inp.id = '__dndProbe'; inp.style.display = 'none'; document.body.appendChild(inp); true;`);
    win.webContents.debugger.attach('1.3');
    const { root } = await win.webContents.debugger.sendCommand('DOM.getDocument');
    const { nodeId } = await win.webContents.debugger.sendCommand('DOM.querySelector', { nodeId: root.nodeId, selector: '#__dndProbe' });
    await win.webContents.debugger.sendCommand('DOM.setFileInputFiles', { nodeId, files: [fx('source.xlsx')] });
    win.webContents.debugger.detach();
    await js(`const f = document.getElementById('__dndProbe').files[0]; const ev = new Event('drop', { bubbles: true, cancelable: true }); Object.defineProperty(ev, 'dataTransfer', { value: { files: [f] } }); document.getElementById('dropA').dispatchEvent(ev); true;`);
    await sleep(300);
    const sheetA = await js(`({ name: document.getElementById('fileNameA').textContent, visible: !document.getElementById('sheetRowA').hidden, options: Array.from(document.getElementById('sheetSelectA').options).map(o => o.value), label: document.querySelector('#sheetRowA label').textContent })`);
    check('드래그앤드롭으로 원본(A) 선택 + 시트 2개 선택 UI', sheetA.name === 'source.xlsx' && sheetA.visible && JSON.stringify(sheetA.options) === '["직원","기타"]' && sheetA.label === '시트 선택', sheetA);

    await pick('dropB', fx('target.xlsx'));
    await click('btnAnalyze');
    let mapping = await mappingOf('mappingTableBody', 'bcolumn');
    const expectedMap = { 성명: '이름', 사번: '사번', 입사릴: '입사일', 급여: '급여', 비율: '비율', 우편번호: '우편번호', 비고: '비고', 비고_1: '비고_1' };
    check('B2에서 시작하는 표 + 헤더 공백 + 중복 컬럼 + 동의어/오타 자동 매핑', JSON.stringify(mapping) === JSON.stringify(expectedMap), mapping);
    const log1 = await logText();
    check('제목 행 건너뛰기 로그 (타겟)', log1.includes('타겟(B) 파일: 헤더 위 1개 행은 건너뜀'), log1.split('\n').filter((l) => l.includes('타겟(B)')).join(' / '));
    check('중복 컬럼 구분 로그', log1.includes('비고_1(으)로 구분했습니다'));
    const previews = await js(`Object.fromEntries(Array.from(document.querySelectorAll('#mappingTableBody tr')).map(tr => [tr.querySelector('select').dataset.bcolumn, tr.querySelector('.mapping-preview-cell').textContent]))`);
    check('미리보기: 날짜는 YYYY-MM-DD(한국 시간대에서도), 앞자리 0 보존, 중복 컬럼은 각자 값', previews['입사릴'] === '2026-01-15 / 2026-02-19 12:00:00' && previews['우편번호'] === '03187 / 04321' && previews['비고'] === '첫째' && previews['비고_1'] === '둘째 / 셋째', previews);

    // 시트 변경 -> 초기화
    await js(`const s = document.getElementById('sheetSelectA'); s.value = '기타'; s.dispatchEvent(new Event('change')); true;`);
    await sleep(200);
    const afterSwitch = await js(`({ status: document.getElementById('analyzeStatus').textContent, rows: document.querySelectorAll('#mappingTableBody select').length })`);
    check('시트 변경 시 매핑 초기화', afterSwitch.status === '대기 중' && afterSwitch.rows === 0, afterSwitch);
    await js(`const s = document.getElementById('sheetSelectA'); s.value = '직원'; s.dispatchEvent(new Event('change')); true;`);
    await click('btnAnalyze');

    // 필수 컬럼 + 변환 확인 모달 취소 -> 저장 안 됨
    await js(`const cb = document.querySelector('#mappingTableBody input[data-bcolumn="비고"]'); cb.checked = true; cb.dispatchEvent(new Event('change')); true;`);
    const outTab1 = path.join(OUT, 'converted.xlsx');
    if (fs.existsSync(outTab1)) fs.unlinkSync(outTab1);
    await click('btnConvert', 300);
    const review = await js(`({ visible: !document.getElementById('convertReviewModal').hidden, summary: document.getElementById('convertReviewSummary').textContent, head: Array.from(document.querySelectorAll('#convertReviewTableHead th')).map(t => t.textContent), firstRow: Array.from(document.querySelectorAll('#convertReviewTableBody tr:first-child td')).map(t => t.textContent) })`);
    check('확인 모달: 필수 컬럼 누락 + 중복 행 감지', review.visible && review.summary.includes('필수 칸 "비고"이(가) 비어 있는 행 1행') && review.summary.includes('모든 값이 같은 중복 행 1행: 5행(3행과 같음)'), review.summary);
    check('확인 모달: 원래 컬럼 이름 + 날짜 표기', JSON.stringify(review.head) === '["성명","사번","입사릴","급여","비율","우편번호","비고","비고"]' && review.firstRow[2] === '2026-01-15', review);
    await click('convertReviewCancel');
    check('확인 모달 취소 시 저장 안 함', !fs.existsSync(outTab1) && (await logText()).includes('확인 단계에서 취소'));

    // 저장 대화상자 취소
    await click('btnConvert', 300);
    await click('convertReviewProceed', 300);
    check('저장 대화상자 취소 처리', !fs.existsSync(outTab1) && (await logText()).includes('저장이 취소되었습니다'));

    // 실제 변환
    q.save.push(outTab1);
    await click('btnConvert', 300);
    await click('convertReviewProceed', 100);
    // 저장은 임시 파일에 쓰고 다시 열어 확인(별도 프로세스)한 뒤 원래 이름으로 바꾸므로 파일이 생길 때까지 기다린다
    for (let i = 0; i < 150 && !fs.existsSync(outTab1); i++) await sleep(100);
    await sleep(100);
    if (!fs.existsSync(outTab1)) {
      check('변환 파일 저장', false, await lastAlert());
    } else {
      const c = readOut(outTab1);
      const header = c[0].map((x) => x && x.v);
      check('결과: 헤더는 원래 이름(중복 비고 포함)', JSON.stringify(header) === '["성명","사번","입사릴","급여","비율","우편번호","비고","비고"]', header);
      const r1 = c[1];
      check('결과: 텍스트로 된 사번은 텍스트 유지', r1[1].t === 's' && r1[1].v === '20230001', r1[1]);
      check('결과: 날짜는 엑셀 날짜 + yyyy-mm-dd (하루 밀림 없음)', r1[2].t === 'n' && r1[2].v === 46037 && /yyyy-mm-dd/.test(r1[2].z || ''), r1[2]);
      check('결과: 시각 있는 날짜는 시각까지', c[2][2].v === 46072.5 && /hh:mm:ss/.test(c[2][2].z || ''), c[2][2]);
      check('결과: 급여는 숫자 + 천단위', r1[3].t === 'n' && r1[3].v === 1500000 && r1[3].z === '#,##0', r1[3]);
      check('결과: 비율은 원본 형식(0%) 유지', r1[4].t === 'n' && r1[4].v === 0.15 && r1[4].z === '0%', r1[4]);
      check('결과: 우편번호 앞자리 0 보존', r1[5].v === '03187', r1[5]);
      check('결과: 중복 컬럼이 각자 값', r1[6].v === '첫째' && r1[7].v === '둘째' && c[2][7].v === '셋째', [r1[6], r1[7], c[2][7]]);
    }

    // 매핑 규칙 JSON 저장 -> 매핑 망가뜨리기 -> 불러오기
    const ruleFile = path.join(OUT, 'rule.json');
    q.save.push(ruleFile);
    await click('btnSaveMapping');
    const rule = fs.existsSync(ruleFile) ? JSON.parse(fs.readFileSync(ruleFile, 'utf-8')) : null;
    check('매핑 규칙 JSON 저장 (필수 컬럼 포함)', rule && rule.mapping['성명'] === '이름' && JSON.stringify(rule.requiredColumns) === '["비고"]', rule);
    await js(`document.querySelectorAll('#mappingTableBody select').forEach(s => { s.value = ''; s.dispatchEvent(new Event('change')); }); true;`);
    q.open.push([ruleFile]);
    await click('btnLoadMapping');
    mapping = await mappingOf('mappingTableBody', 'bcolumn');
    const reqChecked = await js(`document.querySelector('#mappingTableBody input[data-bcolumn="비고"]').checked`);
    const prevAfterLoad = await js(`document.querySelector('#mappingTableBody select[data-bcolumn="성명"]').closest('tr').querySelector('.mapping-preview-cell').textContent`);
    check('매핑 규칙 불러오기 (매핑 + 필수 + 미리보기 갱신)', JSON.stringify(mapping) === JSON.stringify(expectedMap) && reqChecked && prevAfterLoad === '홍길동 / 김영희', { mapping, reqChecked, prevAfterLoad });

    // 프리셋: 저장 -> 망가뜨리기 -> 적용 -> 같은 이름 덮어쓰기 -> 두 번째 저장 -> 삭제
    const savePreset = async (name) => {
      await js(`document.getElementById('btnSavePreset').click(); true;`);
      await sleep(150);
      await js(`document.getElementById('presetNameInput').value = ${JSON.stringify(name)}; document.getElementById('presetModalSave').click(); true;`);
      await sleep(200);
    };
    await savePreset('직원명부');
    await js(`document.querySelectorAll('#mappingTableBody select').forEach(s => { s.value = ''; s.dispatchEvent(new Event('change')); }); true;`);
    const presetId = await js(`Array.from(document.getElementById('presetSelect').options).find(o => o.textContent === '직원명부').value`);
    await js(`const p = document.getElementById('presetSelect'); p.value = ${JSON.stringify(presetId)}; p.dispatchEvent(new Event('change')); true;`);
    await sleep(200);
    mapping = await mappingOf('mappingTableBody', 'bcolumn');
    check('프리셋 저장 후 적용 시 매핑 복원', JSON.stringify(mapping) === JSON.stringify(expectedMap), mapping);
    await js(`const p = document.getElementById('presetSelect'); p.value = ''; p.dispatchEvent(new Event('change')); true;`);
    await savePreset('직원명부');
    const overwriteAsked = await js(`window.__lastConfirm || ''`);
    const sameNameCount = await js(`Array.from(document.getElementById('presetSelect').options).filter(o => o.textContent === '직원명부').length`);
    check('같은 이름 프리셋은 덮어쓰기 확인 후 하나만 유지', overwriteAsked.includes('이미 "직원명부"') && sameNameCount === 1, { overwriteAsked, sameNameCount });
    await savePreset('삭제할프리셋');
    const bothExist = await js(`['직원명부', '삭제할프리셋'].every(n => Array.from(document.getElementById('presetSelect').options).some(o => o.textContent === n))`);
    check('선택된 프리셋이 있을 때 다른 이름으로 저장하면 새 프리셋 (원래 것 유지)', bothExist);
    await js(`const p = document.getElementById('presetSelect'); p.value = Array.from(p.options).find(o => o.textContent === '삭제할프리셋').value; p.dispatchEvent(new Event('change')); true;`);
    await sleep(150);
    await click('btnDeletePreset');
    const names = await js(`Array.from(document.getElementById('presetSelect').options).map(o => o.textContent)`);
    check('프리셋 삭제', !names.includes('삭제할프리셋') && names.includes('직원명부'), names);

    // .xls 타겟
    await pick('dropB', fx('target.xls'));
    await click('btnAnalyze');
    mapping = await mappingOf('mappingTableBody', 'bcolumn');
    check('.xls(구형 엑셀) 타겟 분석', mapping['성명'] === '이름' && mapping['비고_1'] === '비고_1', mapping);

    // CP949 CSV -> 앞자리 0/날짜 글자 보존
    await pick('dropA', fx('source_cp949.csv'));
    await pick('dropB', fx('target_csv.xlsx'));
    await click('btnAnalyze');
    mapping = await mappingOf('mappingTableBody', 'bcolumn');
    check('CP949 CSV 한글 헤더 인식 + 자동 매핑', JSON.stringify(mapping) === JSON.stringify({ 성명: '이름', 연락처: '전화번호', 우편번호: '우편번호', 가입일: '가입일' }), mapping);
    const outCsv = path.join(OUT, 'converted_csv.xlsx');
    q.save.push(outCsv);
    await click('btnConvert', 300);
    await click('convertReviewProceed', 100);
    // 저장은 임시 파일에 쓰고 다시 열어 확인(별도 프로세스)한 뒤 원래 이름으로 바꾸므로 파일이 생길 때까지 기다린다
    for (let i = 0; i < 150 && !fs.existsSync(outCsv); i++) await sleep(100);
    await sleep(100);
    if (fs.existsSync(outCsv)) {
      const r = readOut(outCsv)[1].map((x) => x && x.v);
      // 가입일은 CSV 글자 "2026-01-15"를 날짜로 읽어 엑셀 날짜(일련번호 46037, yyyy-mm-dd)로 저장한다
      const joined = readOut(outCsv)[1][3];
      check('CSV 변환: 전화번호/우편번호 앞자리 0은 글자, 가입일은 날짜', JSON.stringify(r) === '["홍길동","01012345678","03187",46037]' && joined && joined.z === 'yyyy-mm-dd', { r, joined });
    } else check('CSV 변환 저장', false, await lastAlert());
    await pick('dropA', fx('source_bom.csv'));
    await click('btnAnalyze');
    const bomPreview = await js(`document.querySelector('#mappingTableBody select[data-bcolumn="연락처"]').closest('tr').querySelector('.mapping-preview-cell').textContent`);
    check('UTF-8 BOM CSV', bomPreview === '01055556666', bomPreview);

    // 로그 지우기
    await click('btnClearLog', 100);
    check('로그 지우기', (await logText()) === '');

    // 동의어 사전: 추가 저장 (2단계에서 유지 확인)
    await click('btnManageSynonymsTab1', 150);
    await js(`const t = document.getElementById('synonymGroupsTextarea'); t.value = t.value + '\\n급여, 월급, 봉급'; document.getElementById('synonymManagerSave').click(); true;`);
    await sleep(200);
    const syn = await js(`require(${JSON.stringify(path.join(REPO, 'common.js'))}).findAutoMatch('월급', ['급여', '이름'], new Set())`);
    check('동의어 사전 추가가 즉시 자동 매핑에 반영', syn.match === '급여' && syn.reason === 'synonym', syn);

    /* ================= TAB 2 ================= */
    await tab('tab2');
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

    // 탭2 동의어 버튼도 같은 모달
    await click('btnManageSynonymsTab2', 150);
    check('탭2 동의어 사전 버튼', await js(`!document.getElementById('synonymManagerModal').hidden`));
    await click('synonymManagerCancel', 100);
  }

  if (PHASE === '2') {
    await sleep(300);
    const names = await js(`Array.from(document.getElementById('presetSelect').options).map(o => o.textContent)`);
    check('재시작 후 프리셋 유지', names.includes('직원명부') && !names.includes('삭제할프리셋'), names);
    await click('btnManageSynonymsTab1', 150);
    check('재시작 후 동의어 사전 유지', await js(`document.getElementById('synonymGroupsTextarea').value.includes('급여, 월급, 봉급')`));
    // 기본값 복원 버튼
    await click('synonymManagerReset', 100);
    const resetText = await js(`document.getElementById('synonymGroupsTextarea').value`);
    check('동의어 사전 기본값 되돌리기', !resetText.includes('월급') && resetText.includes('이름, 성명, 성함'));
    await click('synonymManagerCancel', 100);
    // 프리셋 적용 후 변환까지 (재시작 후에도 동작)
    await pick('dropA', fx('source.xlsx'));
    await pick('dropB', fx('target.xlsx'));
    await click('btnAnalyze');
    await js(`document.querySelectorAll('#mappingTableBody select').forEach(s => { s.value = ''; s.dispatchEvent(new Event('change')); }); const p = document.getElementById('presetSelect'); p.value = Array.from(p.options).find(o => o.textContent === '직원명부').value; p.dispatchEvent(new Event('change')); true;`);
    await sleep(200);
    const m = await mappingOf('mappingTableBody', 'bcolumn');
    check('재시작 후 프리셋 적용', m['성명'] === '이름' && m['비고_1'] === '비고_1', m);
  }

  if (PHASE === '3') {
    await sleep(300);
    const log3 = await logText();
    check('손상된 프리셋 파일 → 경고 후 빈 목록으로 계속', log3.includes('프리셋을 불러오는 중 오류') && (await js(`document.getElementById('presetSelect').options.length`)) === 1, log3);
    check('손상된 동의어 파일 → 경고 후 기본 사전으로 계속', log3.includes('동의어 사전을 불러오는 중 오류'), log3);
    const syn = await js(`require(${JSON.stringify(path.join(REPO, 'common.js'))}).findAutoMatch('성명', ['이름'], new Set())`);
    check('손상 복구 후에도 기본 동의어 매핑 동작', syn.match === '이름', syn);
  }

  check('렌더러 오류 로그 없음', consoleErrors.length === 0, consoleErrors);
  const failed = results.filter((r) => !r).length;
  flog(`=== phase ${PHASE}: ${results.length - failed}/${results.length} passed ===`);
  app.exit(failed ? 1 : 0);
});
