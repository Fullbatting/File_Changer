// tests/e2e/many-columns.test.js
// 컬럼이 많을 때(원본 52열 × 타겟 41열, 120행): 상태별 묶음 · 검색 · 차례로 확인 · 추천 · 프리셋 제안 · 열 지도.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/many-columns.test.js

const fs = require('fs');
const path = require('path');
const { createHarness } = require('./harness.js');
const { makeMany, MANY_TARGET } = require('./fixtures.js');

const h = createHarness('many-columns');
const { js, sleep, click, record, answerConfirm } = h;
const q = (s) => JSON.stringify(s);

const rowsShown = () => js(`Array.from(document.querySelectorAll('#wbRows .wb-row')).map((r) => r.dataset.b)`);
const heads = () => js(`Array.from(document.querySelectorAll('#wbRows .wb-group-head')).map((g) => [g.dataset.group, g.querySelector('.wb-count').textContent, g.classList.contains('collapsed')])`);
const chip = (id) => js(`document.getElementById(${q(id)}).textContent`);
const active = () => js(`document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.b || null : null`);
const press = (code, keyName) => js(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: ${q(keyName)}, code: ${q(code)}, bubbles: true, cancelable: true })); true;`);
const setSearch = (id, value) => js(`const s = document.getElementById(${q(id)}); s.value = ${q(value)}; s.dispatchEvent(new Event('input')); true;`);

function seedPreset() {
  const mapping = {};
  const auto = { 성명: '이름', 사번: '사번', 생일: '생년월일', 이메일: '이메일주소', 소속: '부서', 직위: '직급', 입사일자: '입사일', 직책수당: '직급수당', 주민번호: '주민등록번호', 계좌번호: '급여계좌번호', 연락처: '휴대폰' };
  MANY_TARGET.slice(0, 38).forEach((b) => {
    mapping[b] = auto[b] || b;
  });
  const preset = { id: 'p1', name: '급여대장 10월', sourceHeaders: [], targetHeaders: MANY_TARGET.slice(0, 38), mapping, requiredColumns: ['연락처'], createdAt: '', updatedAt: '' };
  fs.writeFileSync(path.join(h.userDataDir, 'mapping-presets.json'), JSON.stringify([preset]));
}

async function run() {
  seedPreset();
  await h.start();
  const fx = makeMany(h.workDir, 120);
  await h.loadFiles(fx.source, fx.target);

  /* 20) 프리셋 제안 띠 */
  const banner = await js(`({ shown: !document.getElementById('wbBanner').hidden, text: document.getElementById('wbBanner').textContent })`);
  record('비슷한 프리셋 제안 띠: "41열 중 38열"', banner.shown && banner.text.includes('급여대장 10월') && banner.text.includes('41열 중 38열'), banner);
  record('제안만 하고 자동 적용하지 않음(연락처는 아직 연결 없음)', (await js(`document.querySelector('#mappingTableBody select[data-bcolumn="연락처"]').value`)) === '', '');

  /* 16) 상태별 묶음 */
  let hs = await heads();
  record('타겟 41열 → 기본 "상태별": 확인 필요 4 · 연결 없음 4 · 자동 연결됨 33(접힘)', JSON.stringify(hs) === JSON.stringify([['fuzzy', '4', false], ['none', '4', false], ['auto', '33', true]]), hs);
  record('요약 칩: 연결 37/41 · 확인 필요 4 · 연결 없음 4', (await chip('wbStatConnected')) === '연결 37 / 41' && (await chip('wbStatFuzzy')) === '확인 필요 4' && (await chip('wbStatNone')) === '연결 없음 4', await chip('wbStatConnected'));
  let rows = await rowsShown();
  record('보이는 줄: 확인 필요 4 + 연결 없음 4', JSON.stringify(rows) === JSON.stringify(['입사일자', '직책수당', '주민번호', '계좌번호', '연락처', '국민연금', '주소', '비고']), rows);
  const ambiguous = await js(`({ text: document.querySelector('.wb-row[data-b="주소"] .wb-a-empty').textContent, chips: Array.from(document.querySelectorAll('.wb-row[data-b="주소"] .wb-cand')).map((c) => c.textContent) })`);
  record('모호해서 자동 연결 안 한 "주소": 후보 2개(주소1·주소2) 표시', ambiguous.text.includes('후보 2개가 비슷함') && JSON.stringify(ambiguous.chips) === JSON.stringify(['주소1', '주소2']), ambiguous);
  const tray = await js(`({ count: document.querySelector('#wbTray .wb-count').textContent, footer: document.querySelector('.wb-tray-footer').textContent })`);
  record('트레이: 연결 안 된 원본 9개, 값이 모두 빈 컬럼 6개 숨김', tray.count === '9' && tray.footer.includes('6개 숨김'), tray);
  await js(`document.querySelector('[data-group-toggle="auto"]').click(); true;`);
  await sleep(100);
  hs = await heads();
  record('접힌 "자동 연결됨"을 펼치면 줄이 보임', hs[2][2] === false && (await rowsShown()).includes('성명'), hs);
  await js(`document.querySelector('[data-group-toggle="auto"]').click(); true;`);
  await sleep(100);

  /* 18) 차례로 확인하기: J/K/Y, 남은 N개 모두 맞음 */
  await js(`document.querySelector('.wb-row[data-b="입사일자"]').focus(); true;`);
  await press('KeyJ', 'j');
  await sleep(50);
  const afterJ = await active();
  await press('KeyK', 'k');
  await sleep(50);
  const afterK = await active();
  record('J/K: 다음·이전 확인 항목으로 초점 이동', afterJ === '직책수당' && afterK === '입사일자', { afterJ, afterK });
  await press('KeyY', 'y');
  await sleep(150);
  const afterY = await active();
  hs = await heads();
  record('Y: 맞음 처리 후 "직접 연결" 묶음으로 옮겨지고 초점은 다음 확인 항목으로', afterY === '직책수당' && hs.some(([g, n]) => g === 'manual' && n === '1') && (await chip('wbStatFuzzy')) === '확인 필요 3', { afterY, hs });
  await press('KeyY', 'ㅛ'); // 한글 자판에서도(키 위치 기준) 동작
  await sleep(150);
  record('한글 자판 상태에서도 Y(ㅛ) 동작', (await chip('wbStatFuzzy')) === '확인 필요 2', await chip('wbStatFuzzy'));
  await click('wbUndo', 120);
  await click('wbConfirmAll', 150);
  const dlg = await answerConfirm(true);
  await sleep(150);
  record('[남은 N개 모두 맞음]: 확인 창에 연결 목록, 확인 후 확인 필요 0', dlg.open && dlg.items.length === 3 && dlg.items.includes('직책수당 ← 직급수당') && (await chip('wbStatFuzzy')) === '확인 필요 0', dlg);
  await click('wbUndo', 150);
  record('일괄 맞음은 실행 취소 한 번으로 모두 되돌림', (await chip('wbStatFuzzy')) === '확인 필요 3', await chip('wbStatFuzzy'));

  /* 17) 검색: 이름 · 초성 · 값 */
  await setSearch('wbSearch', 'ㅇㄹㅊ');
  await sleep(80);
  rows = await rowsShown();
  record('연결 보기 초성 검색 "ㅇㄹㅊ" → 연락처', JSON.stringify(rows) === JSON.stringify(['연락처']), rows);
  await setSearch('wbSearch', '@corp');
  await sleep(80);
  rows = await rowsShown();
  record('값 검색 "@corp" → 메일 값이 연결된 이메일 줄(접힌 묶음 안이어도 보임)', JSON.stringify(rows) === JSON.stringify(['이메일']), rows);
  await setSearch('wbSearch', '');
  await sleep(80);

  /* 19) 추천 · 고르기 창의 값 검색 */
  await js(`document.querySelector('.wb-row[data-b="연락처"] .wb-b-card').click(); true;`);
  await sleep(100);
  const pick = await js(`({ labels: Array.from(document.querySelectorAll('#wbPickerList .wb-picker-label')).map((l) => l.firstChild.textContent), recs: Array.from(document.querySelectorAll('#wbPickerList .wb-pick-item')).slice(0, 2).map((b) => [b.dataset.a, Array.from(b.querySelectorAll('.wb-reason')).map((r) => r.textContent)]) })`);
  record('연락처 추천 1·2순위 모두 전화번호 값(비상연락처·휴대폰)', pick.labels[0] === '추천' && pick.recs.every(([, reasons]) => reasons.includes('값: 전화번호')) && pick.recs.map(([a]) => a).sort().join() === '비상연락처,휴대폰', pick);
  const expanders = await js(`Array.from(document.querySelectorAll('#wbPickerList .wb-pick-expand')).map((b) => b.textContent)`);
  record('고르기 창: "이미 연결된 컬럼 37"·"값이 모두 빈 컬럼 6" 접힘', JSON.stringify(expanders) === JSON.stringify(['이미 연결된 컬럼 37', '값이 모두 빈 컬럼 6']), expanders);
  await setSearch('wbPickerSearch', '@corp');
  await sleep(80);
  const byValue = await js(`Array.from(document.querySelectorAll('#wbPickerList .wb-pick-item')).map((b) => b.dataset.a).filter(Boolean)`);
  record('고르기 창 값 검색 "@corp" → 이메일주소·E-mail', JSON.stringify(byValue) === JSON.stringify(['이메일주소', 'E-mail']), byValue);
  await js(`document.getElementById('wbPickerSearch').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); true;`);
  await sleep(80);

  /* 연결 없음 칩 → 첫 연결 없음 줄로 */
  await click('wbStatNone', 120);
  record('요약 "연결 없음" 칩 → 첫 연결 없음 줄로 초점', (await active()) === '연락처', await active());

  /* 15) 가상 스크롤: "양식 순서"에서 보이는 줄만 DOM */
  await click('wbGroupOrder', 120);
  rows = await rowsShown();
  record('양식 순서: 41줄 중 보이는 줄만 DOM에 있음', rows.length > 5 && rows.length < 41 && rows[0] === '성명', rows.length);
  await js(`const s = document.getElementById('wbRows'); s.scrollTop = s.scrollHeight; s.dispatchEvent(new Event('scroll')); true;`);
  await sleep(150);
  rows = await rowsShown();
  record('끝까지 스크롤하면 마지막 줄(비고)이 그려짐', rows.includes('비고') && !rows.includes('성명'), rows.slice(-3));
  await click('wbGroupStatus', 120);

  /* 21) 결과 미리보기: 열 지도 · 열 고정 · 가로 가상화 · 문제 있는 열만 · 페이지 */
  const map = await js(`({ shown: !document.getElementById('wbColMap').hidden, cells: document.querySelectorAll('#wbColMap .wb-colmap-cell').length, contact: document.querySelector('#wbColMap [data-jump="연락처"]').className, fuzzy: document.querySelector('#wbColMap [data-jump="직책수당"]').className, pinned: document.querySelector('#wbColMap [data-jump="성명"]').className })`);
  record('열 지도: 41칸, 상태 색(연결 없음·확인 필요), 고정 열 표시', map.shown && map.cells === 41 && /status-none/.test(map.contact) && /status-fuzzy/.test(map.fuzzy) && /pinned/.test(map.pinned), map);
  let ths = await js(`Array.from(document.querySelectorAll('#wbGrid thead th[data-b]')).map((th) => th.dataset.b)`);
  record('가로 가상화: 41열 중 보이는 열만 그림, 첫 열(성명) 고정', ths.length < 41 && ths[0] === '성명' && (await js(`document.querySelector('#wbGrid thead th[data-b="성명"]').classList.contains('sticky')`)), ths.length);
  await js(`document.querySelector('#wbColMap [data-jump="비고"]').click(); true;`);
  await sleep(200);
  ths = await js(`Array.from(document.querySelectorAll('#wbGrid thead th[data-b]')).map((th) => th.dataset.b)`);
  const scrollLeft = await js(`document.getElementById('wbGridWrap').scrollLeft`);
  record('열 지도 칸을 누르면 그 열로 가로 스크롤, 고정 열은 남아 있음', scrollLeft > 0 && ths.includes('비고') && ths[0] === '성명' && !ths.includes('사번'), { scrollLeft, ths });
  const inView = await js(`document.querySelector('#wbColMap [data-jump="비고"]').classList.contains('in-view')`);
  record('열 지도에 지금 보이는 열 표시', inView, '');
  await click('wbColProblem', 150);
  ths = await js(`Array.from(document.querySelectorAll('#wbGrid thead th[data-b]')).map((th) => th.dataset.b)`);
  // 앞에서 입사일자를 [맞음]으로 확인했으므로 문제 열은 확인 필요 3 + 연결 없음 4
  record('문제 있는 열만 7 → 고정 열 + 문제 열 7개', (await chip('wbColProblem')) === '문제 있는 열만 7' && ths.length === 8 && ths[0] === '성명' && !ths.includes('입사일자'), ths);
  await click('wbColAll', 120);
  await js(`document.querySelector('#wbGrid thead th[data-b="성명"] [data-pin]').click(); true;`);
  await sleep(100);
  record('핀으로 고정 풀기', !(await js(`document.querySelector('#wbGrid thead th[data-b="성명"]') && document.querySelector('#wbGrid thead th[data-b="성명"]').classList.contains('sticky')`)), '');
  let range = await js(`document.getElementById('wbPreviewRange').textContent`);
  record('페이지: 전체 120행 중 1–50', range.startsWith('전체 120행 중 1–50'), range);
  await click('wbNextPage', 120);
  range = await js(`document.getElementById('wbPreviewRange').textContent`);
  const firstRowNum = await js(`document.querySelector('#wbGrid tbody td.wb-rownum').textContent`);
  record('다음 50행: 51–100, 첫 행은 시트 52행', range.startsWith('전체 120행 중 51–100') && firstRowNum === '52', { range, firstRowNum });

  /* 연결 보기에서 줄을 고르면 미리보기가 그 열로 */
  await js(`document.querySelector('.wb-row[data-b="계좌번호"]').focus(); true;`);
  await sleep(150);
  const linked = await js(`({ th: !!document.querySelector('#wbGrid thead th[data-b="계좌번호"].wb-hl') })`);
  record('연결 보기 줄 선택 → 미리보기 그 열 강조·스크롤', linked.th, linked);

  /* 20) 프리셋 적용 → 실행 취소 */
  await click('wbApplyPreset', 200);
  const applied = await js(`({ a: document.querySelector('#mappingTableBody select[data-bcolumn="연락처"]').value, badge: document.querySelector('.wb-row[data-b="연락처"] .wb-badge') ? document.querySelector('.wb-row[data-b="연락처"] .wb-badge').textContent : '(접힘)', banner: document.getElementById('wbBanner').hidden, required: document.querySelector('#mappingTableBody input[data-bcolumn="연락처"]').checked })`);
  record('[프리셋 적용]: 연락처←휴대폰(이유 프리셋)·필수 반영, 띠 사라짐', applied.a === '휴대폰' && applied.banner && applied.required, applied);
  await click('wbUndo', 150);
  record('프리셋 적용도 실행 취소 한 단계', (await js(`document.querySelector('#mappingTableBody select[data-bcolumn="연락처"]').value`)) === '', '');
}

run()
  .then(() => h.finish())
  .catch((err) => h.fail(err));
