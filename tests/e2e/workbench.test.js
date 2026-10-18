// tests/e2e/workbench.test.js
// 매핑 작업대 시나리오 테스트(작은 예제: 원본 9열 × 타겟 6열).
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/workbench.test.js

const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { createHarness } = require('./harness.js');
const { makeSmall, serial } = require('./fixtures.js');

const h = createHarness('workbench');
const { js, sleep, click, record, answerConfirm, queues } = h;

const q = (s) => JSON.stringify(s);
const listMapping = () => js(`Object.fromEntries(Array.from(document.querySelectorAll('#mappingTableBody select')).map((s) => [s.dataset.bcolumn, s.value]))`);
const rowInfo = (b) =>
  js(`const r = document.querySelector('.wb-row[data-b=' + ${q(JSON.stringify(b))} + ']');
    r ? { a: r.querySelector('.wb-a-card b') ? r.querySelector('.wb-a-card b').textContent : '', badge: r.querySelector('.wb-b-card .wb-badge').textContent, line: r.querySelector('.wb-line').className, dashed: getComputedStyle(r.querySelector('.wb-line')).getPropertyValue('--line').trim() } : null`);
const chip = (id) => js(`document.getElementById(${q(id)}).textContent`);
const cellText = (i, b) => js(`const td = document.querySelector('td.wb-cell[data-r="${i}"][data-b=' + ${q(JSON.stringify(b))} + ']'); td ? td.textContent : null`);
const cellClass = (i, b) => js(`const td = document.querySelector('td.wb-cell[data-r="${i}"][data-b=' + ${q(JSON.stringify(b))} + ']'); td ? td.className : null`);
const key = (selector, keyName, code, extra = '') =>
  js(`const t = ${selector}; t.dispatchEvent(new KeyboardEvent('keydown', { key: ${q(keyName)}, code: ${q(code || '')}, bubbles: true, cancelable: true ${extra} })); true;`);

/** 미리보기 칸을 더블클릭해 값을 넣고 keyName(Enter/Escape/Tab)으로 끝낸다 */
async function editCell(i, b, value, keyName = 'Enter') {
  await js(`const td = document.querySelector('td.wb-cell[data-r="${i}"][data-b=' + ${q(JSON.stringify(b))} + ']'); td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); true;`);
  await sleep(50);
  await js(`const input = document.getElementById('wbCellInput'); input.value = ${q(value)}; true;`);
  await key(`document.getElementById('wbCellInput')`, keyName, keyName === 'Tab' ? 'Tab' : keyName);
  await sleep(150);
}

/** 원본 카드(트레이)를 타겟 카드로 끌어다 놓는다 — 실제 좌표로 pointer 이벤트를 보낸다 */
async function dragTo(fromSelector, toSelector) {
  return js(`
    const from = ${fromSelector};
    const to = ${toSelector};
    from.scrollIntoView({ block: 'center' });
    to.scrollIntoView({ block: 'nearest' });
    const a = from.getBoundingClientRect();
    const b = to.getBoundingClientRect();
    const ax = a.left + 30, ay = a.top + a.height / 2;
    const bx = b.left + b.width / 2, by = b.top + b.height / 2;
    const opt = (x, y) => ({ bubbles: true, cancelable: true, clientX: x, clientY: y, button: 0, pointerId: 1, isPrimary: true });
    from.dispatchEvent(new PointerEvent('pointerdown', opt(ax, ay)));
    window.dispatchEvent(new PointerEvent('pointermove', opt(ax + 20, ay)));
    window.dispatchEvent(new PointerEvent('pointermove', opt(bx, by)));
    const overlay = !!document.querySelector('.wb-drag-overlay');
    const highlighted = to.classList.contains('wb-drop-target') || !!to.closest('.wb-drop-target');
    window.dispatchEvent(new PointerEvent('pointerup', opt(bx, by)));
    ({ overlay, highlighted });
  `);
}

async function run() {
  await h.start();
  const fx = makeSmall(h.workDir);

  /* 1) 분석 직후: 자동 연결이 이유별 선·배지로 보이고 요약 칩 수가 맞다 */
  await h.loadFiles(fx.source, fx.target);
  const status = await js(`document.getElementById('analyzeStatus').textContent`);
  record('구조 분석 완료', status === '분석 완료', status);
  let mapping = await listMapping();
  record('자동 매핑: 성명←이름, 사번, 입사릴←입사일, 급여, 부서명←부서, 연락처 없음', JSON.stringify(mapping) === JSON.stringify({ 성명: '이름', 사번: '사번', 입사릴: '입사일', 급여: '급여', 부서명: '부서', 연락처: '' }), mapping);
  const reasons = {};
  for (const b of ['성명', '사번', '입사릴', '연락처']) reasons[b] = await rowInfo(b);
  record('연결 이유 배지: 동의어/같은 이름/비슷한 이름/연결 없음', reasons.성명.badge === '동의어' && reasons.사번.badge === '같은 이름' && reasons.입사릴.badge === '비슷한 이름' && reasons.연락처.badge === '연결 없음', reasons);
  record('비슷한 이름 연결은 점선(주황) 선', /reason-fuzzy/.test(reasons.입사릴.line) && /reason-exact/.test(reasons.사번.line), reasons.입사릴.line);
  const chips = { c: await chip('wbStatConnected'), f: await chip('wbStatFuzzy'), n: await chip('wbStatNone'), d: await chip('wbStatDuplicate'), e: await chip('wbStatEdited') };
  record('요약 칩: 연결 5/6 · 확인 필요 1 · 연결 없음 1 · 중복 1행 · 고친 값 0개', chips.c === '연결 5 / 6' && chips.f === '확인 필요 1' && chips.n === '연결 없음 1' && chips.d === '중복 1행' && chips.e === '고친 값 0개', chips);
  const rowNums = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  record('미리보기 원본 행 번호 = 실제 시트 행(B2 시작 · 제목 행 있는 표)', JSON.stringify(rowNums) === JSON.stringify(['4', '5', '6', '7중복']), rowNums);
  const dupTitle = await js(`document.querySelector('#wbGrid tbody tr:nth-child(4) td.wb-rownum').title`);
  record('중복 행에 "5행과 모든 값이 같음" 표시', dupTitle === '5행과 모든 값이 같음', dupTitle);
  const previewTexts = { name: await cellText(0, '성명'), date: await cellText(0, '입사릴'), money: await cellText(0, '급여'), id: await cellText(0, '사번') };
  record('미리보기 표시 = 엑셀 표시(날짜 yyyy-mm-dd, 급여 천 단위, 사번 글자 그대로)', previewTexts.name === '홍길동' && previewTexts.date === '2026-01-15' && previewTexts.money === '1,500,000' && previewTexts.id === '20230001', previewTexts);
  const tray = await js(`({ items: Array.from(document.querySelectorAll('#wbTrayList .wb-tray-item b')).map((b) => b.textContent), footer: document.querySelector('.wb-tray-footer').textContent })`);
  record('연결 안 된 원본 컬럼 트레이: 휴대폰·우편번호·비고, 빈 컬럼 1개 숨김', JSON.stringify(tray.items) === JSON.stringify(['휴대폰', '우편번호', '비고']) && tray.footer.includes('1개 숨김'), tray);
  const chipsInEmpty = await js(`Array.from(document.querySelectorAll('.wb-row[data-b="연락처"] .wb-cand')).map((c) => c.textContent)`);
  record('연결 없는 줄에 추천 후보 칩(휴대폰: 값 모양이 전화번호)', chipsInEmpty[0] === '휴대폰', chipsInEmpty);

  /* 2) 끌어다 놓기 · 팝오버 · 키보드로 연결 */
  const drag = await dragTo(`document.querySelector('#wbTrayList [data-drag-a="휴대폰"]')`, `document.querySelector('.wb-row[data-b="연락처"] .wb-b-card')`);
  await sleep(150);
  let info = await rowInfo('연락처');
  record('끌어다 놓기: 휴대폰 → 연락처 연결, 이유 "직접 연결"', info && info.a === '휴대폰' && info.badge === '직접 연결', { drag, info });
  record('끄는 동안 임시 선과 놓을 곳 강조가 보인다', drag.overlay && drag.highlighted, drag);
  record('미리보기 연락처 열이 즉시 채워짐', (await cellText(0, '연락처')) === '010-1234-5678', await cellText(0, '연락처'));
  await click('wbUndo', 150);
  info = await rowInfo('연락처');
  record('실행 취소: 연락처 연결이 풀림', info.badge === '연결 없음', info);

  await js(`document.querySelector('.wb-row[data-b="연락처"] .wb-b-card').click(); true;`);
  await sleep(100);
  const picker = await js(`({ open: !!document.getElementById('wbPicker'), first: document.querySelector('#wbPickerList .wb-pick-item') ? document.querySelector('#wbPickerList .wb-pick-item').dataset.a : null, reasons: Array.from(document.querySelectorAll('#wbPickerList .wb-pick-item:first-of-type .wb-reason')).map((r) => r.textContent), focus: document.activeElement && document.activeElement.id })`);
  record('타겟 카드를 누르면 고르기 창, 추천 1순위 휴대폰(값: 전화번호), 검색창에 초점', picker.open && picker.first === '휴대폰' && picker.reasons.includes('값: 전화번호') && picker.focus === 'wbPickerSearch', picker);
  await key(`document.getElementById('wbPickerSearch')`, 'Enter', 'Enter');
  await sleep(150);
  info = await rowInfo('연락처');
  record('Enter로 추천 연결', info.a === '휴대폰' && !(await js(`!!document.getElementById('wbPicker')`)), info);
  await click('wbUndo', 150);

  // 키보드만으로: 줄에 초점 → C → 초성 검색 → Enter
  await js(`document.querySelector('.wb-row[data-b="연락처"]').focus(); true;`);
  await key(`document.activeElement`, 'c', 'KeyC');
  await sleep(100);
  await js(`const s = document.getElementById('wbPickerSearch'); s.value = 'ㅎㄷㅍ'; s.dispatchEvent(new Event('input')); true;`);
  await sleep(80);
  const found = await js(`Array.from(document.querySelectorAll('#wbPickerList .wb-pick-item')).map((b) => b.dataset.a).filter(Boolean)`);
  record('고르기 창 초성 검색: "ㅎㄷㅍ" → 휴대폰', JSON.stringify(found) === JSON.stringify(['휴대폰']), found);
  await key(`document.getElementById('wbPickerSearch')`, 'ArrowDown', 'ArrowDown');
  await key(`document.getElementById('wbPickerSearch')`, 'ArrowUp', 'ArrowUp');
  await key(`document.getElementById('wbPickerSearch')`, 'Enter', 'Enter');
  await sleep(150);
  info = await rowInfo('연락처');
  const focusAfter = await js(`document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.b : null`);
  record('키보드(C·검색·Enter)로 연결, 초점은 그 줄로 돌아옴', info.a === '휴대폰' && focusAfter === '연락처', { info, focusAfter });
  await js(`document.querySelector('.wb-row[data-b="연락처"] .wb-b-card').click(); true;`);
  await sleep(80);
  await key(`document.getElementById('wbPickerSearch')`, 'Escape', 'Escape');
  await sleep(80);
  record('Esc로 고르기 창 닫기', !(await js(`!!document.getElementById('wbPicker')`)), '');

  /* 3) [맞음] → 점선이 실선, 확인 필요 감소 */
  await js(`document.querySelector('.wb-row[data-b="입사릴"] [data-action="confirm"]').click(); true;`);
  await sleep(150);
  info = await rowInfo('입사릴');
  record('[맞음]: 이유 "확인됨", 선이 실선, 확인 필요 0', info.badge === '확인됨' && !/reason-fuzzy/.test(info.line) && (await chip('wbStatFuzzy')) === '확인 필요 0', info);

  /* 4) 값 고치기 · 필수 빈칸 · 원래 값으로 */
  await js(`document.querySelector('.wb-row[data-b="연락처"] [data-action="required"]').click(); true;`);
  await sleep(120);
  record('필수 지정 → 필수 빈칸 1행(6행 연락처 비어 있음)', (await chip('wbStatMissing')) === '필수 빈칸 1행' && /missing/.test(await cellClass(2, '연락처')), await chip('wbStatMissing'));
  await editCell(2, '연락처', '010-5555-1234');
  record('셀 수정(Enter): 노란 칸, 고친 값 1개, 필수 빈칸 0행', /edited/.test(await cellClass(2, '연락처')) && (await cellText(2, '연락처')) === '010-5555-1234' && (await chip('wbStatEdited')) === '고친 값 1개' && (await chip('wbStatMissing')) === '필수 빈칸 0행', { cls: await cellClass(2, '연락처'), e: await chip('wbStatEdited') });
  const toast = await js(`({ shown: !document.getElementById('wbToast').hidden, text: document.getElementById('wbToast').textContent })`);
  record('고친 뒤 "실행 취소" 알림', toast.shown && toast.text.includes('6행 연락처'), toast);
  await editCell(0, '성명', '바뀔리없음', 'Escape');
  record('Esc로 취소하면 값이 바뀌지 않음', (await cellText(0, '성명')) === '홍길동' && (await chip('wbStatEdited')) === '고친 값 1개', await cellText(0, '성명'));
  await editCell(0, '급여', '협의', 'Tab');
  const activeAfterTab = await js(`const a = document.activeElement; a && a.dataset ? [a.dataset.r, a.dataset.b] : null`);
  record('Tab: 확정 후 오른쪽 칸(부서명)으로 이동', JSON.stringify(activeAfterTab) === JSON.stringify(['0', '부서명']) && (await cellText(0, '급여')) === '협의', activeAfterTab);

  await js(`document.querySelector('td.wb-cell[data-r="2"][data-b="연락처"]').click(); true;`);
  await sleep(80);
  const bubble = await js(`({ shown: !document.getElementById('wbBubble').hidden, text: document.getElementById('wbBubble').textContent })`);
  record('고친 칸을 누르면 "원래 값: (빈칸)" 말풍선', bubble.shown && bubble.text.includes('원래 값: (빈칸)'), bubble);
  await click('wbRevertCell', 150);
  record('[원래 값으로]: 고친 값이 지워지고 다시 빈칸', (await cellText(2, '연락처')) === '' && (await chip('wbStatEdited')) === '고친 값 1개', { t: await cellText(2, '연락처'), e: await chip('wbStatEdited') });
  await click('wbUndo', 150);
  record('실행 취소로 고친 값 복원', (await cellText(2, '연락처')) === '010-5555-1234', await cellText(2, '연락처'));

  /* 5) 중복이던 행을 고치면 중복이 사라진다 */
  await editCell(1, '성명', '김영희(본사)');
  record('중복 행의 짝을 고치면 중복 0행', (await chip('wbStatDuplicate')) === '중복 0행', await chip('wbStatDuplicate'));

  /* 6) 고친 값이 있는 열의 연결 변경 → 확인 → 그 열의 고친 값만 지워짐 → 실행 취소로 복원 */
  await js(`document.querySelector('.wb-row[data-b="연락처"] [data-action="disconnect"]').click(); true;`);
  await sleep(150);
  let dlg = await answerConfirm(false);
  record('고친 값이 있는 열의 연결 끊기 → 확인 창("고친 값 1개가 지워집니다")', dlg.open && dlg.message.includes('연락처 열에서 직접 고친 값 1개가 지워집니다'), dlg);
  record('확인 창에서 취소하면 아무것도 바뀌지 않음', (await rowInfo('연락처')).a === '휴대폰' && (await chip('wbStatEdited')) === '고친 값 3개', await chip('wbStatEdited'));
  await js(`document.querySelector('.wb-row[data-b="연락처"] [data-action="disconnect"]').click(); true;`);
  await sleep(150);
  await answerConfirm(true);
  record('확인하면 그 열의 고친 값만 지워짐(성명·급여 고친 값은 남음)', (await rowInfo('연락처')).badge === '연결 없음' && (await chip('wbStatEdited')) === '고친 값 2개', await chip('wbStatEdited'));
  await click('wbUndo', 150);
  record('실행 취소: 연결과 고친 값 모두 복원', (await rowInfo('연락처')).a === '휴대폰' && (await chip('wbStatEdited')) === '고친 값 3개' && (await cellText(2, '연락처')) === '010-5555-1234', await chip('wbStatEdited'));

  /* 7) 실행 취소 / 다시 실행 */
  await click('wbUndo', 120);
  record('실행 취소: 성명 고친 값이 풀려 중복 1행으로', (await chip('wbStatDuplicate')) === '중복 1행', await chip('wbStatDuplicate'));
  await click('wbRedo', 120);
  record('다시 실행: 중복 0행', (await chip('wbStatDuplicate')) === '중복 0행', await chip('wbStatDuplicate'));
  await js(`document.body.focus(); document.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', code: 'KeyZ', ctrlKey: true, bubbles: true })); true;`);
  await sleep(120);
  record('Ctrl+Z 단축키', (await chip('wbStatDuplicate')) === '중복 1행', await chip('wbStatDuplicate'));
  await js(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'y', code: 'KeyY', ctrlKey: true, bubbles: true })); true;`);
  await sleep(120);
  record('Ctrl+Y 단축키', (await chip('wbStatDuplicate')) === '중복 0행', await chip('wbStatDuplicate'));

  /* 8) 저장 파일 = 화면의 미리보기 (숫자·날짜로 고친 값은 그 종류로, 열 판정은 그대로) */
  await editCell(2, '급여', '3,000,000');
  await editCell(1, '입사릴', '2026-03-01');
  const previewGrid = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr')).map((tr) => Array.from(tr.querySelectorAll('td.wb-cell')).map((td) => td.textContent))`);
  const outPath = path.join(h.workDir, 'out.xlsx');
  queues.save.push(outPath);
  await click('btnConvert', 300);
  const review = await js(`({ open: !document.getElementById('convertReviewModal').hidden, text: document.getElementById('convertReviewSummary').textContent })`);
  record('변환 확인 창: 요약(총 4행 · 직접 고친 값 포함)', review.open && review.text.includes('총 4행') && review.text.includes('직접 고친 값 5개 포함'), review);
  await click('convertReviewProceed', 900);
  const saved = fs.existsSync(outPath);
  record('결과 파일 저장', saved, outPath);
  if (saved) {
    const wb = XLSX.readFile(outPath, { cellNF: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const header = [0, 1, 2, 3, 4, 5].map((c) => (ws[XLSX.utils.encode_cell({ r: 0, c })] || {}).v);
    record('저장 파일 헤더 = 타겟 양식', JSON.stringify(header) === JSON.stringify(['성명', '사번', '입사릴', '급여', '부서명', '연락처']), header);
    const cellAt = (r, c) => ws[XLSX.utils.encode_cell({ r, c })] || {};
    const savedGrid = [1, 2, 3, 4].map((r) => [0, 1, 2, 3, 4, 5].map((c) => {
      const cell = cellAt(r, c);
      if (cell.v === undefined || cell.v === '') return '';
      return cell.t === 'n' ? XLSX.SSF.format(cell.z || 'General', cell.v) : String(cell.v);
    }));
    record('저장 파일의 보이는 값 = 미리보기 표(모든 칸)', JSON.stringify(savedGrid) === JSON.stringify(previewGrid), { savedGrid, previewGrid });
    const salaryText = cellAt(1, 3);
    const salaryNum = cellAt(3, 3);
    record('숫자 열에 글자로 고친 칸("협의")만 글자, 나머지는 숫자', salaryText.t === 's' && salaryText.v === '협의' && cellAt(2, 3).t === 'n' && cellAt(2, 3).z === '#,##0', { salaryText, other: cellAt(2, 3) });
    record('숫자 열에 "3,000,000"으로 고친 칸은 숫자 3000000', salaryNum.t === 'n' && salaryNum.v === 3000000, salaryNum);
    const date = cellAt(2, 2);
    record('날짜 열에 "2026-03-01"로 고친 칸은 날짜(yyyy-mm-dd)', date.t === 'n' && date.v === serial(2026, 3, 1) && date.z === 'yyyy-mm-dd', date);
    record('사번은 원본처럼 글자로 저장', cellAt(1, 1).t === 's' && cellAt(1, 1).v === '20230001', cellAt(1, 1));
  }

  /* 9) 문제 있는 행만 필터 */
  await editCell(3, '연락처', '');
  await click('wbRowFilter', 120);
  let filtered = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  record('문제 있는 행만: 고친 행·빈칸 행만 보임', filtered.length === 4 && filtered.some((t) => t.startsWith('7')), filtered);
  await editCell(3, '연락처', '010-7777-8888');
  filtered = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  record('필터 중 문제가 풀린 행은 "해결"로 남아 보임', filtered.some((t) => t === '7해결'), filtered);
  await click('wbRowFilter', 120);
  await click('wbStatMissing', 120);
  record('요약의 "필수 빈칸" 칩을 누르면 문제 있는 행만 켜짐', await js(`document.getElementById('wbRowFilter').getAttribute('aria-pressed') === 'true'`), '');
  await click('wbRowFilter', 120);

  /* 12) 목록 보기 ↔ 연결 보기 */
  await click('wbViewList', 150);
  const listShown = await js(`({ list: !document.getElementById('wbListView').hidden, conn: !document.getElementById('wbConnect').hidden })`);
  mapping = await listMapping();
  record('목록 보기 전환: 같은 연결 상태', listShown.list && !listShown.conn && mapping.연락처 === '휴대폰' && mapping.입사릴 === '입사일', { listShown, mapping });
  await js(`const s = document.querySelector('#mappingTableBody select[data-bcolumn="부서명"]'); s.value = '비고'; s.dispatchEvent(new Event('change')); true;`);
  await sleep(150);
  await click('wbViewConnect', 150);
  info = await rowInfo('부서명');
  record('목록 보기에서 바꾼 연결이 연결 보기에도 반영(직접 연결)', info.a === '비고' && info.badge === '직접 연결', info);
  await click('wbUndo', 120);

  /* 10) 프리셋·규칙 파일 저장/불러오기 — 이전 형식 호환, 고친 값은 저장하지 않음 */
  const rulePath = path.join(h.workDir, 'rule.json');
  queues.save.push(rulePath);
  await click('btnSaveMapping', 200);
  const rule = JSON.parse(fs.readFileSync(rulePath, 'utf-8'));
  record('매핑 규칙 파일: 연결·필수만 저장(고친 값 없음)', rule.mapping.연락처 === '휴대폰' && rule.requiredColumns.includes('연락처') && !JSON.stringify(rule).includes('010-5555'), Object.keys(rule));
  const oldRule = path.join(h.workDir, 'old_rule.json');
  fs.writeFileSync(oldRule, JSON.stringify({ 성명: '이름', 연락처: '우편번호', 사번: '없는컬럼' }));
  queues.open.push([oldRule]);
  await click('btnLoadMapping', 200);
  await answerConfirm(true); // 연락처 열에 고친 값이 있으므로 확인
  await sleep(100);
  mapping = await listMapping();
  info = await rowInfo('연락처');
  record('단순 {B: A} 형식 규칙 파일 불러오기(없는 컬럼은 비움), 이유 "매핑 규칙"', mapping.연락처 === '우편번호' && mapping.사번 === '' && info.badge === '매핑 규칙', { mapping, info });
  await click('wbUndo', 120);
  // 프리셋 저장
  await click('btnSavePreset', 120);
  await js(`document.getElementById('presetNameInput').value = '직원명부'; document.getElementById('presetModalSave').click(); true;`);
  await sleep(200);
  const presetsFile = path.join(h.userDataDir, 'mapping-presets.json');
  const presets = JSON.parse(fs.readFileSync(presetsFile, 'utf-8'));
  record('프리셋 저장: 기존 형식(mapping·requiredColumns), 고친 값 없음', presets.length === 1 && presets[0].mapping.연락처 === '휴대폰' && Array.isArray(presets[0].requiredColumns) && !JSON.stringify(presets).includes('010-5555'), presets[0] && Object.keys(presets[0]));

  /* 11) 파일·시트 변경 / 재분석 시 확인 후 초기화 */
  const editedBefore = await chip('wbStatEdited');
  queues.open.push([fx.target]);
  await click('dropA', 200);
  dlg = await answerConfirm(false);
  const fileLabel = await js(`document.getElementById('fileNameA').textContent`);
  record('고친 값이 있을 때 원본 파일을 바꾸면 확인 창, 취소하면 그대로', dlg.open && dlg.title === '원본 파일을 바꿀까요?' && fileLabel === '직원명부.xlsx' && (await chip('wbStatEdited')) === editedBefore, { dlg, fileLabel });
  await click('btnAnalyze', 200);
  dlg = await answerConfirm(true);
  await sleep(300);
  record('재분석은 확인 후 상태 전체 초기화', dlg.open && dlg.title === '다시 분석할까요?' && (await chip('wbStatEdited')) === '고친 값 0개' && (await js(`document.getElementById('wbUndo').disabled`)), dlg.title);

  const alerts = await js(`window.__alerts`);
  record('알림창은 변환 완료 안내 하나뿐', JSON.stringify(alerts) === JSON.stringify(['✅ 변환이 완료되었습니다.']), alerts);
}

run()
  .then(() => h.finish())
  .catch((err) => {
    console.log(`ERROR ${err && err.stack ? err.stack : err}`);
    h.finish();
  });
