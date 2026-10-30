// tests/unit/output-check.test.js
// 엑셀 쓰기(excel-writer) → 다시 열기(SheetJS) → 비교(output-check) 단위 테스트: node tests/unit/output-check.test.js

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const writer = require('../../excel-writer.js');
const check = require('../../output-check.js');

let passed = 0;
let failed = 0;
const queue = [];
function test(title, fn) {
  queue.push({ title, fn });
}

const headers = ['이름', '코드', '급여', '비율', '입사일', '일시', '확인', '비고'];
const labels = ['성명', '코드', '급여', '비율', '입사일', '일시', '확인', '비고'];
const rows = [
  { 이름: '홍길동', 코드: '00123', 급여: 1500000, 비율: 0.15, 입사일: new Date(2026, 0, 15), 일시: new Date(2026, 0, 15, 13, 45, 10), 확인: true, 비고: '  앞뒤 공백 ' },
  { 이름: '김영희', 코드: '04321', 급여: 2250000.5, 비율: 0.2, 입사일: new Date(2026, 1, 19), 일시: '', 확인: false, 비고: '' },
  { 이름: '이철수', 코드: '', 급여: '협의', 비율: 1, 입사일: '', 일시: new Date(1999, 11, 31, 23, 59, 59), 확인: '', 비고: '=1+1' },
];
const spec = (sheets) => ({ headers, labels, sheets, numberFormats: { 비율: '0.0%' }, columnKinds: headers.map((h) => ({ kind: ['급여', '비율'].includes(h) ? 'number' : 'text' })) });
const verifySpec = (sheets) => ({
  headers,
  labels,
  sheets,
  formatOf: writer.makeCellFormatter(headers, rows, spec(sheets)),
  toSerial: writer.toExcelDateSerial,
});
const reopen = (buffer) => XLSX.read(buffer, { type: 'buffer', cellNF: true });

test('저장 → 다시 열기: 글자(앞자리 0·공백·=로 시작)·숫자·형식·날짜·시각·참거짓·빈 칸이 모두 같음', async () => {
  const sheets = [{ name: '명부', rows }];
  const result = check.verifyWorkbook(reopen(await writer.writeStyledExcel(spec(sheets))), verifySpec(sheets));
  assert.deepStrictEqual(result.problems.map(check.describeProblem), []);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.checkedCells, rows.length * headers.length);
  assert.deepStrictEqual(result.rowCounts, { 명부: 3 });
});

test('스트리밍으로 파일에 써도 같고, 진행률을 알린다', async () => {
  const many = Array.from({ length: 4500 }, (_, i) => ({ ...rows[i % 3], 이름: `직원${i}` }));
  const sheets = [{ name: '명부', rows: many }];
  const file = path.join(os.tmpdir(), `oc-${process.pid}.xlsx`);
  const progress = [];
  await writer.writeStyledExcel({ ...spec(sheets), onProgress: (d, t) => progress.push([d, t]) }, file);
  const result = check.verifyWorkbook(XLSX.readFile(file, { cellNF: true }), { ...verifySpec(sheets), formatOf: writer.makeCellFormatter(headers, many, spec(sheets)) });
  fs.unlinkSync(file);
  assert.strictEqual(result.ok, true, result.problems.map(check.describeProblem).join('\n'));
  assert.deepStrictEqual(progress, [[2000, 4500], [4000, 4500], [4500, 4500]]);
});

test('취소하면 CancelledError', async () => {
  const many = Array.from({ length: 4500 }, () => rows[0]);
  await assert.rejects(writer.writeStyledExcel({ ...spec([{ name: 'S', rows: many }]), shouldCancel: () => true }), (err) => err.cancelled === true);
});

test('dense 모드로 다시 열어도 같은 결과', async () => {
  const sheets = [{ name: '명부', rows }];
  const buffer = await writer.writeStyledExcel(spec(sheets));
  const result = check.verifyWorkbook(XLSX.read(buffer, { type: 'buffer', cellNF: true, dense: true }), verifySpec(sheets));
  assert.strictEqual(result.ok, true, result.problems.map(check.describeProblem).join('\n'));
});

test('시트 여러 개: 이름·순서·시트별 행 수', async () => {
  const sheets = [{ name: '전체', rows }, { name: '영업팀', rows: rows.slice(0, 1) }, { name: '(빈 값)', rows: rows.slice(1) }];
  const result = check.verifyWorkbook(reopen(await writer.writeStyledExcel(spec(sheets))), verifySpec(sheets));
  assert.strictEqual(result.ok, true, result.problems.map(check.describeProblem).join('\n'));
  assert.deepStrictEqual(result.rowCounts, { 전체: 3, 영업팀: 1, '(빈 값)': 2 });
});

test('다른 점을 찾아낸다: 글자→숫자로 바뀐 코드, 날짜 어긋남, 형식, 헤더, 빠진 행, 시트 이름', async () => {
  const sheets = [{ name: '명부', rows }];
  const wb = reopen(await writer.writeStyledExcel(spec(sheets)));
  const ws = wb.Sheets['명부'];
  ws.B2 = { t: 'n', v: 123 }; // 00123이 숫자 123으로
  ws.E3 = { ...ws.E3, v: ws.E3.v - 1 }; // 날짜 하루 앞
  ws.D2 = { ...ws.D2, z: '0%' }; // 형식
  ws.A1 = { t: 's', v: '이름' }; // 헤더
  const r = check.verifyWorkbook(wb, verifySpec(sheets));
  const reasons = r.problems.map((p) => `${p.rowNumber}:${p.column}:${p.reason}`);
  assert.deepStrictEqual(reasons, ['1:성명:헤더 이름이 다름', '2:코드:글자가 다름', '2:비율:숫자 형식이 다름', '3:입사일:날짜가 다름']);
  assert.ok(check.describeProblem(r.problems[1]).includes('저장하려던 값 00123(글자) / 파일 123(숫자)'), check.describeProblem(r.problems[1]));

  delete ws.A4;
  ws['!ref'] = 'A1:H3';
  wb.SheetNames = ['결과'];
  const r2 = check.verifyWorkbook(wb, verifySpec(sheets));
  assert.ok(r2.problems.some((p) => p.kind === 'sheets'));
});

test('끝쪽 행이 모두 빈 칸이어도 행 수 차이로 보지 않는다', async () => {
  const blankTail = [...rows, Object.fromEntries(headers.map((h) => [h, '']))];
  const sheets = [{ name: 'S', rows: blankTail }];
  const result = check.verifyWorkbook(reopen(await writer.writeStyledExcel(spec(sheets))), verifySpec(sheets));
  assert.strictEqual(result.ok, true, result.problems.map(check.describeProblem).join('\n'));
});

(async () => {
  for (const { title, fn } of queue) {
    try {
      await fn();
      passed += 1;
      console.log(`PASS - ${title}`);
    } catch (err) {
      failed += 1;
      console.log(`FAIL - ${title} :: ${err.message}`);
    }
  }
  console.log(`\n[output-check] ${passed}/${passed + failed} 통과`);
  process.exit(failed ? 1 : 0);
})();
