// tests/e2e/merge-split.test.js
// 원본 합치기(기준 컬럼 자동 추천·왼쪽/양쪽·보고)와 결과 시트 나누기(추천·저장)를 확인한다.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/merge-split.test.js

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { createHarness } = require('./harness.js');
const { sheetFromRows, writeBook, moneyCell } = require('./fixtures.js');

const h = createHarness('merge-split');
const { js, sleep, click, record, queues } = h;
const extras = () =>
  js(`Array.from(document.querySelectorAll('#extraSourceList .extra-source')).map((it) => ({
    name: it.querySelector('.extra-source-title b').textContent,
    keys: Array.from(it.querySelectorAll('.extra-source-keys select')).map((s) => s.value),
    hint: it.querySelector('.extra-source-hint') ? it.querySelector('.extra-source-hint').textContent : '',
    counts: it.querySelector('.extra-match-counts') ? it.querySelector('.extra-match-counts').textContent : '',
    zero: it.querySelector('.extra-zero input').checked,
  }))`);
const listMapping = () => js(`Object.fromEntries(Array.from(document.querySelectorAll('#mappingTableBody select')).map((s) => [s.dataset.bcolumn, s.value]))`);

async function run() {
  await h.start();
  const dir = h.workDir;
  const staff = writeBook(path.join(dir, '직원명부.xlsx'), {
    직원: sheetFromRows([
      ['사번', '이름', '부서'],
      ['20230001', '홍길동', '영업팀'],
      ['20230002', '김영희', '영업팀'],
      ['20230003', '이철수', '개발팀'],
      ['20230004', '박민수', '인사팀'],
    ]),
  });
  const payCsv = path.join(dir, '급여.csv');
  fs.writeFileSync(payCsv, '﻿직원번호,이름,기본급,수당\r\n20230002,김영희,2250000,100000\r\n20230001,홍길동,1500000,50000\r\n20230009,최지은,1800000,0\r\n20230001,홍길동,9999999,9\r\n');
  const contact = writeBook(path.join(dir, '연락처.xlsx'), {
    연락처: sheetFromRows([
      ['사번', '휴대폰번호'],
      ['20230003', '010-3333-0003'],
      ['20230001', '010-1111-0001'],
    ]),
  });
  const target = writeBook(path.join(dir, '급여대장.xlsx'), { 명부: sheetFromRows([['성명', '사번', '부서', '기본급', '수당', '연락처']]) });

  queues.open.push([staff]);
  await click('dropA', 250);
  queues.open.push([target]);
  await click('dropB', 250);

  /* 기준 컬럼 자동 추천 */
  queues.open.push([payCsv]);
  await click('btnAddSource', 300);
  queues.open.push([contact]);
  await click('btnAddSource', 300);
  let ex = await extras();
  record('추가 원본 2개가 목록에 보임', ex.length === 2 && ex[0].name === '급여.csv' && ex[1].name === '연락처.xlsx', ex);
  record('이름이 다른 기준 컬럼도 값으로 추천(사번 = 직원번호, CSV 숫자 ↔ 엑셀 글자)', JSON.stringify(ex[0].keys) === JSON.stringify(['사번', '직원번호', 'left', 'first']) && ex[0].hint.startsWith('추천: 값 일치'), ex[0]);
  record('같은 이름 기준 컬럼 추천(사번 = 사번)', ex[1].keys[0] === '사번' && ex[1].keys[1] === '사번' && ex[1].hint.includes('이름 같음'), ex[1]);
  record('기본은 정확 비교(앞자리 0 무시 꺼짐) + 연결될 행 수 미리 표시', !ex[0].zero && ex[0].counts === '정확히 비교하면 원본(A) 2행 연결', ex[0]);
  await js(`const b = document.querySelectorAll('#extraSourceList .extra-zero input')[0]; b.checked = true; b.dispatchEvent(new Event('change')); true;`);
  await sleep(300);
  ex = await extras();
  record('앞자리 0 무시를 켜면 무시할 때의 연결 행 수도 함께 표시', ex[0].zero && ex[0].counts === '정확히 비교하면 원본(A) 2행 연결 · 앞자리 0을 무시하면 2행 연결', ex[0]);
  await js(`const b = document.querySelectorAll('#extraSourceList .extra-zero input')[0]; b.checked = false; b.dispatchEvent(new Event('change')); true;`);
  await sleep(300);

  /* 합쳐서 분석 */
  await click('btnAnalyze', 600);
  const status = await js(`document.getElementById('analyzeStatus').textContent`);
  const report = await js(`Array.from(document.querySelectorAll('#mergeReport .merge-report-line')).map((l) => l.textContent)`);
  const details = await js(`Array.from(document.querySelectorAll('#mergeReport .merge-report-detail')).map((l) => l.textContent)`);
  record('합쳐서 분석 완료 + 합치기 결과 보고 2줄', status === '분석 완료' && report.length === 2, { status, report });
  record('급여 요약: 연결 2행 · 급여 행 2개 빠짐', report[0] === '급여: 기준 사번 = 직원번호 · 연결된 행 2개 · 급여 행 2개 빠짐', report[0]);
  record(
    '급여 상세: 빠진 행을 이유별 행 번호와 함께(짝 없는 원본 4·5행, 급여에만 있는 4행, 겹친 기준값 5행)',
    details.includes('짝이 없는 원본(A) 행 2개 — 추가 컬럼은 빈칸 (행 4, 5)') &&
      details.includes('급여에만 있는 행 1개는 넣지 않음 (행 4)') &&
      details.includes('급여에서 기준값이 겹쳐 첫 행만 이음 — 나머지 1행은 넣지 않음 (행 5)'),
    details
  );
  let mapping = await listMapping();
  record('합친 컬럼으로 자동 매핑(기본급·수당 ← 급여, 연락처 ← 휴대폰번호)', mapping.기본급 === '기본급' && mapping.수당 === '수당' && mapping.연락처 === '휴대폰번호' && mapping.성명 === '이름', mapping);
  const origin = await js(`const r = document.querySelector('.wb-row[data-b="기본급"] .wb-origin'); r ? r.textContent : ''`);
  record('연결 보기에 합친 컬럼의 출처 표시(급여)', origin === '급여', origin);
  const cells = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr:first-child td.wb-cell')).map((t) => t.textContent)`);
  record('미리보기 첫 행: 홍길동 · 1,500,000(겹친 기준값은 첫 행 사용) · 010-1111-0001', JSON.stringify(cells) === JSON.stringify(['홍길동', '20230001', '영업팀', '1,500,000', '50,000', '010-1111-0001']), cells);
  const third = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr:nth-child(3) td.wb-cell')).map((t) => t.textContent)`);
  record('짝 없는 행(이철수)은 급여 칸이 빈칸, 연락처는 채워짐', third[3] === '' && third[5] === '010-3333-0003', third);

  /* 양쪽 행 모두 */
  await js(`const sel = document.querySelectorAll('#extraSourceList .extra-source')[0].querySelector('select[aria-label="합치는 방식"]'); sel.value = 'outer'; sel.dispatchEvent(new Event('change')); true;`);
  await sleep(300);
  record('합치기 설정을 바꾸면 분석 결과 초기화', (await js(`document.getElementById('analyzeStatus').textContent`)) === '대기 중', '');
  await click('btnAnalyze', 600);
  const rowNums = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  const last = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr:last-child td.wb-cell')).map((t) => t.textContent)`);
  const fifth = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr:nth-child(5) td.wb-cell')).map((t) => t.textContent)`);
  record(
    '양쪽 행 모두: 급여에만 있던 최지은(급여 4)과 겹친 기준값의 두 번째 홍길동(급여 5)까지 빠짐없이 뒤에 추가',
    rowNums.length === 6 && rowNums[4] === '급여 4' && rowNums[5] === '급여 5' && fifth[1] === '20230009' && fifth[3] === '1,800,000' && last[1] === '20230001' && last[3] === '9,999,999',
    { rowNums, fifth, last }
  );

  /* 겹치는 기준값 모두 잇기(왼쪽 기준): 홍길동 행이 급여 두 행만큼 늘어난다 */
  await js(`const it = document.querySelectorAll('#extraSourceList .extra-source')[0]; const m = it.querySelector('select[aria-label="합치는 방식"]'); m.value = 'left'; m.dispatchEvent(new Event('change')); true;`);
  await sleep(300);
  await js(`const it = document.querySelectorAll('#extraSourceList .extra-source')[0]; const d = it.querySelector('select[aria-label="기준값이 겹칠 때"]'); d.value = 'all'; d.dispatchEvent(new Event('change')); true;`);
  await sleep(300);
  await click('btnAnalyze', 600);
  const allNums = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  const pays = await js(`Array.from(document.querySelectorAll('#wbGrid tbody tr td.wb-cell[data-b="기본급"]')).map((t) => t.textContent)`);
  const allDetails = await js(`Array.from(document.querySelectorAll('#mergeReport .merge-report-detail')).map((l) => l.textContent)`);
  record(
    '겹치는 기준값 모두 잇기: 홍길동이 2행(같은 원본 행 2), 늘어난 행 수 보고',
    JSON.stringify(allNums) === JSON.stringify(['2', '2', '3', '4', '5']) && pays[0] === '1,500,000' && pays[1] === '9,999,999' && allDetails.some((t) => t.includes('결과 행 1개 늘어남')),
    { allNums, pays, allDetails }
  );

  /* 결과 시트 나누기 */
  await js(`document.getElementById('splitSelect').dispatchEvent(new Event('focus')); true;`);
  const opts = await js(`({ hint: document.getElementById('splitHint').textContent, recommended: Array.from(document.querySelectorAll('#splitSelect optgroup[label="추천"] option')).map((o) => o.value) })`);
  record('시트 나누기 추천: 부서(값 종류 적음·나누기 좋은 이름)', opts.recommended[0] === '부서' && opts.hint.includes('추천: 부서 컬럼으로 나누기'), opts);
  await js(`const s = document.getElementById('splitSelect'); s.value = '부서'; s.dispatchEvent(new Event('change')); const c = document.getElementById('splitIncludeAll'); c.checked = true; c.dispatchEvent(new Event('change')); true;`);
  await sleep(100);
  record('나누기 선택 시 시트 개수 안내', (await js(`document.getElementById('splitHint').textContent`)).includes('부서 값마다 시트 3개'), await js(`document.getElementById('splitHint').textContent`));
  const out = path.join(dir, 'out.xlsx');
  queues.save.push(out);
  await click('btnConvert', 300);
  const summary = await js(`document.getElementById('convertReviewSummary').textContent`);
  record('확인 창: 시트 나누기 요약 + 합치기 결과', summary.includes('부서 값마다 시트 3개로 나눠 저장(맨 앞에 전체 시트 포함): 영업팀 3행 · 개발팀 1행 · 인사팀 1행') && summary.includes('원본 합치기 — 급여'), summary);
  await click('convertReviewProceed', 900);
  const wb = XLSX.readFile(out, { cellNF: true });
  record('저장: 전체 시트 + 부서별 시트', JSON.stringify(wb.SheetNames) === JSON.stringify(['명부', '영업팀', '개발팀', '인사팀']), wb.SheetNames);
  const rowsOf = (name) => XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true });
  const sales = rowsOf('영업팀');
  record('부서 시트마다 같은 헤더와 그 부서 행만', JSON.stringify(sales[0]) === JSON.stringify(['성명', '사번', '부서', '기본급', '수당', '연락처']) && sales.length === 4 && sales[1][0] === '홍길동' && sales[2][0] === '홍길동' && sales[3][0] === '김영희', sales);
  record('숫자는 시트를 나눠도 숫자 서식(#,##0)', wb.Sheets['영업팀'].D2.t === 'n' && wb.Sheets['영업팀'].D2.z === '#,##0' && wb.Sheets['영업팀'].D2.v === 1500000, wb.Sheets['영업팀'].D2);
  record('전체 시트는 5행 모두', rowsOf('명부').length === 6, rowsOf('명부').length);

  /* 프리셋·규칙 파일에 나누기 설정 저장 */
  const rulePath = path.join(dir, 'rule.json');
  queues.save.push(rulePath);
  await click('btnSaveMapping', 200);
  const rule = JSON.parse(fs.readFileSync(rulePath, 'utf-8'));
  record('매핑 규칙 파일에 시트 나누기 설정 저장', rule.splitBy === '부서' && rule.splitIncludeAll === true, { splitBy: rule.splitBy, all: rule.splitIncludeAll });
  await js(`const s = document.getElementById('splitSelect'); s.value = ''; s.dispatchEvent(new Event('change')); true;`);
  queues.open.push([rulePath]);
  await click('btnLoadMapping', 300);
  record('규칙 파일을 불러오면 나누기 설정도 돌아옴', (await js(`document.getElementById('splitSelect').value`)) === '부서' && (await js(`document.getElementById('splitIncludeAll').checked`)), '');

  /* 추가 원본 빼기 */
  await js(`document.querySelectorAll('#extraSourceList .extra-remove')[1].click(); true;`);
  await sleep(300);
  ex = await extras();
  record('추가 원본 빼기 → 목록에서 사라지고 분석 초기화', ex.length === 1 && (await js(`document.getElementById('analyzeStatus').textContent`)) === '대기 중', ex);
  await click('btnAnalyze', 600);
  mapping = await listMapping();
  record('뺀 뒤 다시 분석하면 그 파일 컬럼은 없음(연락처 연결 없음)', mapping.연락처 === '' && mapping.기본급 === '기본급', mapping);
}

run()
  .then(() => h.finish())
  .catch((err) => h.fail(err));
