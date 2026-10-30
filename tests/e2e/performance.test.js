// tests/e2e/performance.test.js
// 성능: 1만 행 × 20열에서 연결 변경·실행 취소·셀 수정이 100ms 안쪽,
// 원본 300열 × 타겟 200열에서 전체 다시 그리기 200ms 안쪽, 보이는 줄·열만 DOM에 있는지.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/performance.test.js

const { createHarness } = require('./harness.js');
const { makeWide } = require('./fixtures.js');

const h = createHarness('performance');
const { js, sleep, click, record, queues } = h;

/**
 * 동작 하나의 시간을 두 가지로 잰다.
 * - work: 스크립트 + 레이아웃(앱 코드가 책임지는 부분). 기준(100ms/200ms)은 이것으로 판정한다.
 * - frame: 다음 화면 갱신(그리기)까지. 이 측정 환경(xvfb, GPU 없음, 공유 CPU)은 그리기가 느리고
 *   편차가 커서, 크게 나빠지지 않았는지만 넉넉한 기준으로 본다.
 * 첫 실행은 1회성 준비 비용이 섞이므로 재지 않고, 여러 번 재 중앙값을 쓴다.
 * after는 다음 측정 전에 상태를 되돌리는 코드(시간에 넣지 않음).
 */
async function measure(code, { times = 5, after = '', warmup = true } = {}) {
  const work = [];
  const frame = [];
  for (let i = warmup ? -1 : 0; i < times; i++) {
    const [w, f] = await js(
      `const t0 = performance.now(); ${code}; document.body.offsetHeight; const t1 = performance.now(); new Promise((r) => requestAnimationFrame(() => setTimeout(() => r([t1 - t0, performance.now() - t0]), 0)))`
    );
    if (i >= 0) {
      work.push(w);
      frame.push(f);
    }
    if (after) await js(`${after}; true;`);
    await sleep(80);
  }
  const median = (list) => [...list].sort((x, y) => x - y)[Math.floor(list.length / 2)];
  const result = { work: median(work), frame: median(frame), maxWork: Math.max(...work) };
  result.toString = () => `작업 중앙값 ${result.work.toFixed(1)}ms(최대 ${result.maxWork.toFixed(1)}) · 화면 갱신까지 ${result.frame.toFixed(1)}ms`;
  return result;
}

/** 앱 코드 기준(work)과 화면 갱신 회귀 기준(frame)을 함께 기록한다 */
function recordTiming(title, m, limit, frameLimit = 250) {
  record(`${title}: ${m} — 작업 ${limit}ms 안쪽`, m.work < limit, `${m}`);
  record(`${title}: 화면 갱신까지 ${frameLimit}ms 안쪽(회귀 확인)`, m.frame < frameLimit, `${m}`);
}

const DISCONNECT = `document.querySelector('.wb-row[data-b="컬럼002"] [data-action="disconnect"]').click()`;
const UNDO = `document.getElementById('wbUndo').click()`;
const REDO = `document.getElementById('wbRedo').click()`;

async function run() {
  await h.start();

  /* 14) 1만 행 × 20열 */
  const big = makeWide(h.workDir, { rowCount: 10000, sourceCols: 20, targetCols: 20, name: 'big' });
  await h.loadFiles(big.source, big.target);
  await sleep(500);
  record('1만 행 × 20열 분석 완료', (await js(`document.getElementById('analyzeStatus').textContent`)) === '분석 완료', '');
  await click('wbGroupOrder', 200);

  recordTiming('1만 행: 연결 변경(끊기)', await measure(DISCONNECT, { after: UNDO }), 100);
  await js(`${DISCONNECT}; true;`);
  await sleep(100);
  recordTiming('1만 행: 실행 취소', await measure(UNDO, { after: REDO }), 100);
  await click('wbUndo', 100);
  recordTiming(
    '1만 행: 셀 수정',
    await measure(`
      const td = document.querySelector('td.wb-cell[data-r="3"][data-b="컬럼003"]');
      td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      const input = document.getElementById('wbCellInput');
      input.value = '고친값' + Math.random();
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))`),
    100
  );
  await sleep(100);
  record('셀 수정이 반영됨', (await js(`document.getElementById('wbStatEdited').textContent`)) === '고친 값 1개', await js(`document.getElementById('wbStatEdited').textContent`));

  /* 파일 읽기(파싱)부터 잰 시간: 앱 로그의 "파일 읽기 Nms" */
  const readLog = await js(`(document.getElementById('logArea').value.match(/원본\\(A\\) 파일 읽기 (\\d+)ms/) || [])[1]`);
  record(`1만 행 × 20열: 파일 읽기(파싱부터) ${readLog}ms — 3초 안쪽`, Number(readLog) > 0 && Number(readLog) < 3000, readLog);

  /* 저장 + ④ 다시 열어 확인: 진행 중에도 화면이 멈추지 않는지(가장 긴 멈춤), 결과가 같은지 */
  const outBig = require('path').join(h.workDir, 'big_out.xlsx');
  queues.save.push(outBig);
  await js(`window.__stall = { last: performance.now(), max: 0 }; window.__stallTimer = setInterval(() => { const now = performance.now(); window.__stall.max = Math.max(window.__stall.max, now - window.__stall.last); window.__stall.last = now; }, 10); true;`);
  const saveStart = Date.now();
  await click('btnConvert', 400);
  await click('convertReviewProceed', 50);
  for (let i = 0; i < 600 && (await js(`document.getElementById('convertResult').hidden`)); i++) await sleep(100);
  const saveMs = Date.now() - saveStart;
  const stall = await js(`clearInterval(window.__stallTimer); window.__stall.max`);
  const resultTitle = await js(`document.getElementById('convertResultTitle').textContent`);
  record(`1만 행 저장 + 다시 열어 확인 ${(saveMs / 1000).toFixed(1)}초 — 결과 같음`, resultTitle.startsWith('✅') && saveMs < 20000, resultTitle);
  record(`저장·확인 중 화면이 가장 오래 멈춘 시간 ${Math.round(stall)}ms — 1초 안쪽(확인은 별도 프로세스)`, stall < 1000, stall);

  /* 저장 취소: 파일을 만들지 않음 */
  const outCancel = require('path').join(h.workDir, 'cancel_out.xlsx');
  queues.save.push(outCancel);
  await click('btnConvert', 400);
  await click('convertReviewProceed', 30);
  await click('btnCancelConvert', 50);
  for (let i = 0; i < 200 && !(await js(`document.getElementById('convertProgress').hidden`)); i++) await sleep(100);
  const files = require('fs').readdirSync(h.workDir);
  record('[취소]하면 결과 파일도 임시 파일도 남지 않음', !files.includes('cancel_out.xlsx') && files.every((f) => !f.includes('.tmp')) && (await js(`document.getElementById('logArea').value.includes('저장을 취소했습니다')`)), files);

  /* 15) 원본 300열 × 타겟 200열 */
  const wide = makeWide(h.workDir, { rowCount: 200, sourceCols: 300, targetCols: 200, name: 'wide' });
  queues.open.push([wide.source]);
  await click('dropA', 300);
  await h.answerConfirm(true); // 고친 값이 있으므로 파일 바꾸기 확인
  await sleep(200);
  queues.open.push([wide.target]);
  await click('dropB', 300);
  const analyzeMs = await js(`const t0 = performance.now(); document.getElementById('btnAnalyze').click(); document.body.offsetHeight; performance.now() - t0`);
  await sleep(300);
  record(`300×200 구조 분석(표 정리+자동 매핑+처음 그리기, 파일 읽기는 위에서 따로 잼) ${analyzeMs.toFixed(0)}ms — 1초 안쪽`, analyzeMs < 1000, analyzeMs);
  await click('wbGroupOrder', 200);
  recordTiming('300×200: 목록 보기로 전환(전체 다시 그리기)', await measure(`document.getElementById('wbViewList').click()`, { after: `document.getElementById('wbViewConnect').click()` }), 200, 400);
  recordTiming('300×200: 연결 보기로 전환(전체 다시 그리기)', await measure(`document.getElementById('wbViewConnect').click()`, { after: `document.getElementById('wbViewList').click()` }), 200, 400);
  await click('wbViewConnect', 100);
  recordTiming('300×200: 연결 끊기', await measure(`document.querySelector('.wb-row[data-b="컬럼001"] [data-action="disconnect"]').click()`, { after: UNDO }), 200, 400);

  const dom = await js(`({ rows: document.querySelectorAll('#wbRows .wb-row').length, ths: document.querySelectorAll('#wbGrid thead th[data-b]').length, options: document.querySelectorAll('#mappingTableBody option').length })`);
  record('보이는 줄·열만 DOM에 있음(연결 보기 < 40줄, 미리보기 < 40열)', dom.rows < 40 && dom.ths < 40, dom);
  record('목록 보기 드롭다운은 열 때 채움(300×200 = 6만 항목을 미리 만들지 않음)', dom.options < 1000, dom.options);
  await js(`const s = document.getElementById('wbRows'); s.scrollTop = s.scrollHeight / 2; s.dispatchEvent(new Event('scroll')); true;`);
  await sleep(150);
  const mid = await js(`Array.from(document.querySelectorAll('#wbRows .wb-row')).map((r) => r.dataset.b)`);
  record('세로 스크롤 후 가운데 줄이 정확히 그려짐(컬럼100 근처)', mid.includes('컬럼100') && !mid.includes('컬럼001'), mid.slice(0, 3));
  await js(`const w = document.getElementById('wbGridWrap'); w.scrollLeft = 150 * 120; w.dispatchEvent(new Event('scroll')); true;`);
  await sleep(200);
  const cols = await js(`Array.from(document.querySelectorAll('#wbGrid thead th[data-b]')).map((th) => th.dataset.b)`);
  record('가로 스크롤 후 그 위치의 열이 그려짐(고정 열 + 컬럼120 근처)', cols[0] === '컬럼001' && cols.includes('컬럼121') && !cols.includes('컬럼050'), cols.slice(0, 4));
}

run()
  .then(() => h.finish())
  .catch((err) => h.fail(err));
