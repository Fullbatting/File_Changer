// tests/unit/merge-sources.test.js
// 원본 합치기·시트 나누기 판단 로직 단위 테스트(Electron 없이 Node로): node tests/unit/merge-sources.test.js

const assert = require('assert');
const m = require('../../merge-sources.js');

let passed = 0;
let failed = 0;
function test(title, fn) {
  try {
    fn();
    passed += 1;
    console.log(`PASS - ${title}`);
  } catch (err) {
    failed += 1;
    console.log(`FAIL - ${title} :: ${err.message}`);
  }
}

const table = (headers, rows, extra = {}) => ({ headers, labels: headers, rows: rows.map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i]]))), ...extra });

const staff = table(
  ['사번', '이름', '부서', '입사일'],
  [
    ['20230001', '홍길동', '영업팀', new Date(2026, 0, 15)],
    ['20230002', '김영희', '영업팀', new Date(2026, 1, 19)],
    ['20230003', '이철수', '개발팀', new Date(2026, 2, 2)],
    ['20230004', '박민수', '인사팀', new Date(2026, 3, 1)],
  ],
  { rowNumbers: [2, 3, 4, 5] }
);
const pay = table(
  ['직원번호', '이름', '기본급', '수당'],
  [
    [20230002, '김영희', 2250000, 100000],
    [20230001, '홍길동', 1500000, 50000],
    [20230009, '최지은', 1800000, 0],
    [20230001, '홍길동', 9999999, 9],
  ],
  { rowNumbers: [2, 3, 4, 5], numberFormats: { 기본급: '#,##0' } }
);

test('normalizeKey: 숫자·숫자 글자·날짜·공백, 앞자리 0은 기본으로 그대로(선택하면 무시)', () => {
  assert.strictEqual(m.normalizeKey(20230001), '20230001');
  assert.strictEqual(m.normalizeKey(' 20230001 '), '20230001');
  assert.notStrictEqual(m.normalizeKey('00123'), m.normalizeKey(123));
  assert.strictEqual(m.normalizeKey('00123', { ignoreLeadingZeros: true }), m.normalizeKey(123, { ignoreLeadingZeros: true }));
  assert.strictEqual(m.normalizeKey('000', { ignoreLeadingZeros: true }), '0');
  assert.strictEqual(m.normalizeKey('0'), '0');
  assert.strictEqual(m.normalizeKey(new Date(2026, 0, 5)), '2026-01-05');
  assert.strictEqual(m.normalizeKey(''), '');
});

test('기준 컬럼 추천: 이름이 달라도(사번↔직원번호) 값으로 찾는다', () => {
  const s = m.suggestJoinKeys(staff, pay);
  assert.ok(s.length > 0, '추천 없음');
  assert.strictEqual(s[0].main, '사번');
  assert.strictEqual(s[0].other, '직원번호');
  assert.ok(s[0].reasons[0].startsWith('값 일치'), s[0].reasons.join());
});

test('기준 컬럼 추천: 겹치는 값이 많은 컬럼(부서처럼 한 값이 여러 행)은 추가 원본 기준으로 고르지 않는다', () => {
  const dept = table(['부서', '부서장'], [['영업팀', 'A'], ['영업팀', 'B'], ['개발팀', 'C']]);
  const s = m.suggestJoinKeys(staff, dept);
  assert.ok(!s.some((x) => x.other === '부서'), JSON.stringify(s));
});

test('기준 컬럼 추천: 이름 같음·동의어 가산', () => {
  const a = table(['이름', '번호'], [['홍', 1], ['김', 2], ['이', 3]]);
  const b = table(['성명', '번호'], [['홍', 1], ['김', 2], ['이', 3]]);
  const s = m.suggestJoinKeys(a, b, { synonymGroups: [['이름', '성명']] });
  const byName = s.find((x) => x.main === '이름' && x.other === '성명');
  const byNo = s.find((x) => x.main === '번호' && x.other === '번호');
  assert.ok(byName.reasons.includes('동의어'));
  assert.ok(byNo.reasons.includes('이름 같음'));
});

test('합치기(왼쪽 기준): 행 순서·개수 유지, 기준값으로 값 연결, 짝 없음은 빈칸', () => {
  const { table: t, reports } = m.joinTables(staff, [{ label: '급여', table: pay, keyMain: '사번', keyOther: '직원번호', mode: 'left' }]);
  assert.deepStrictEqual(t.headers, ['사번', '이름', '부서', '입사일', '이름 (급여)', '기본급', '수당']);
  assert.strictEqual(t.rows.length, 4);
  assert.deepStrictEqual(t.rowNumbers, [2, 3, 4, 5]);
  assert.strictEqual(t.rows[0].기본급, 1500000); // 첫 행(홍길동) 사용
  assert.strictEqual(t.rows[1].기본급, 2250000);
  assert.strictEqual(t.rows[2].기본급, ''); // 이철수: 급여 없음
  assert.strictEqual(t.numberFormats.기본급, '#,##0');
  assert.strictEqual(t.columnOrigin.기본급, '급여');
  const r = reports[0];
  assert.strictEqual(r.matchedRows, 2);
  assert.strictEqual(r.unmatchedRows, 2);
  assert.strictEqual(r.otherOnlyRows, 1); // 최지은
  assert.strictEqual(r.duplicateRows, 1); // 홍길동 두 번
  assert.deepStrictEqual(r.unmatchedRowNumbers, [4, 5]);
  assert.deepStrictEqual(r.otherOnlyRowNumbers, [4]);
  assert.deepStrictEqual(r.skippedDuplicateRowNumbers, [5]);
  assert.strictEqual(r.excludedRows, 2);
  const text = m.describeJoinReport(r);
  assert.ok(text.includes('급여에만 있는 행 1개는 넣지 않음 (행 4)'), text);
  assert.ok(text.includes('나머지 1행은 넣지 않음 (행 5)'), text);
});

test('합치기(양쪽 모두): 짝 없는 행·겹치는 기준값 행 모두 뒤에 붙여 어느 행도 빠지지 않는다', () => {
  const { table: t, reports } = m.joinTables(staff, [{ label: '급여', table: pay, keyMain: '사번', keyOther: '직원번호', mode: 'outer' }]);
  assert.strictEqual(t.rows.length, 6);
  assert.strictEqual(t.rows[4].사번, 20230009);
  assert.strictEqual(t.rows[4]['이름 (급여)'], '최지은');
  assert.strictEqual(t.rows[4].이름, '');
  assert.deepStrictEqual(t.rowNumbers.slice(4), ['급여 4', '급여 5']);
  assert.strictEqual(t.rows[5].기본급, 9999999); // 겹치는 기준값의 두 번째 홍길동 행도 남는다
  assert.strictEqual(reports[0].addedRows, 2);
  assert.strictEqual(reports[0].excludedRows, 0);
  // 급여 파일의 모든 행 값이 결과 어딘가에 있다
  pay.rows.forEach((p) => assert.ok(t.rows.some((r) => r.기본급 === p.기본급), p.기본급));
});

test('합치기(양쪽 모두): 기준값이 빈 추가 원본 행도 덧붙인다, 왼쪽 기준이면 행 번호와 함께 보고', () => {
  const extra = table(['직원번호', '메모'], [['', '번호 없음'], [20230001, '홍']], { rowNumbers: [8, 9] });
  const outer = m.joinTables(staff, [{ label: '메모', table: extra, keyMain: '사번', keyOther: '직원번호', mode: 'outer' }]);
  assert.strictEqual(outer.table.rows.length, 5);
  assert.strictEqual(outer.table.rows[4].메모, '번호 없음');
  assert.strictEqual(outer.table.rowNumbers[4], '메모 8');
  const left = m.joinTables(staff, [{ label: '메모', table: extra, keyMain: '사번', keyOther: '직원번호', mode: 'left' }]);
  assert.deepStrictEqual(left.reports[0].emptyKeyRowNumbers, [8]);
  assert.ok(m.describeJoinReport(left.reports[0]).includes('기준값이 빈 행 1개는 넣지 않음 (행 8)'));
});

test('합치기: 겹치는 기준값을 "모두"로 하면 짝마다 행을 만든다', () => {
  const { table: t, reports } = m.joinTables(staff, [{ label: '급여', table: pay, keyMain: '사번', keyOther: '직원번호', dupPolicy: 'all' }]);
  assert.strictEqual(t.rows.length, 5);
  assert.deepStrictEqual(t.rows.slice(0, 2).map((r) => r.기본급), [1500000, 9999999]);
  assert.deepStrictEqual(t.rowNumbers.slice(0, 2), [2, 2]);
  assert.strictEqual(reports[0].expandedRows, 1);
  assert.strictEqual(reports[0].skippedDuplicateRows, 0);
});

test('앞자리 0: 기본은 정확 비교, 무시하면 연결되고 고르기 전 숫자로 보여 준다', () => {
  const a = table(['코드', '이름'], [['00123', '가'], ['00456', '나'], ['00789', '다']]);
  const b = table(['코드', '금액'], [[123, 10], [456, 20], [789, 30]]);
  const strict = m.joinTables(a, [{ label: 'B', table: b, keyMain: '코드', keyOther: '코드' }]);
  assert.strictEqual(strict.reports[0].matchedRows, 0);
  const loose = m.joinTables(a, [{ label: 'B', table: b, keyMain: '코드', keyOther: '코드', ignoreLeadingZeros: true }]);
  assert.strictEqual(loose.reports[0].matchedRows, 3);
  assert.ok(m.describeJoinReport(loose.reports[0]).includes('앞자리 0 무시'));
  const info = m.analyzeKeyMatch(a, '코드', b, '코드');
  assert.deepStrictEqual([info.exactMatched, info.looseMatched, info.collisions], [0, 3, 0]);
  // 같은 파일 안에 0123과 123이 함께 있으면 무시할 때 서로 다른 코드가 같아진다
  const c = table(['코드', '금액'], [['0123', 1], ['123', 2], ['456', 3]]);
  const info2 = m.analyzeKeyMatch(a, '코드', c, '코드');
  assert.strictEqual(info2.collisions, 1);
  assert.deepStrictEqual(info2.collisionExamples, ['0123 · 123']);
  // 추천은 앞자리 0을 무시해야 이어지는 쌍을 표시만 한다
  const s = m.suggestJoinKeys(a, b);
  assert.strictEqual(s[0].needsIgnoreZeros, true);
  assert.ok(s[0].reasons[0].startsWith('앞자리 0을 무시하면'));
});

test('합치기: 원본을 바꾸지 않는다', () => {
  const before = JSON.stringify(staff.rows);
  m.joinTables(staff, [{ label: '급여', table: pay, keyMain: '사번', keyOther: '직원번호' }]);
  assert.strictEqual(JSON.stringify(staff.rows), before);
});

test('시트 나누기 추천: 부서 같은 컬럼을 고르고 값이 모두 다른 컬럼은 뺀다', () => {
  const rows = [['영업팀', 'a1'], ['영업팀', 'a2'], ['개발팀', 'a3'], ['개발팀', 'a4'], ['인사팀', 'a5'], ['인사팀', 'a6']];
  const s = m.suggestSplitColumns(['부서명', '사번'], rows.length, (i, b) => rows[i][b === '부서명' ? 0 : 1]);
  assert.strictEqual(s.length, 1);
  assert.strictEqual(s[0].column, '부서명');
  assert.ok(s[0].reasons.includes('값 3종류'));
});

test('시트 이름: 금지 문자·31자·대소문자 무시 중복·빈 이름', () => {
  const used = new Set();
  assert.strictEqual(m.sanitizeSheetName('영업/1팀', used), '영업_1팀');
  assert.strictEqual(m.sanitizeSheetName('영업/1팀', used), '영업_1팀 (2)');
  assert.strictEqual(m.sanitizeSheetName('', used), '(빈 값)');
  assert.ok(m.sanitizeSheetName('가'.repeat(40), used).length <= 31);
  assert.strictEqual(m.sanitizeSheetName('History', used), 'History (시트)');
});

test('행 묶기: 처음 나온 순서, 빈 값 묶음', () => {
  const vals = ['B', 'A', 'B', '', 'A'];
  const g = m.groupRowsBy(vals.length, (i) => vals[i]);
  assert.deepStrictEqual(g.map((x) => [x.label, x.indices]), [['B', [0, 2]], ['A', [1, 4]], ['(빈 값)', [3]]]);
});

test('행 묶기: 실제로 "(빈 값)"이라고 적힌 칸과 빈 칸은 다른 묶음, 시트 이름도 따로', () => {
  const vals = ['(빈 값)', '', '(빈 값)', ''];
  const g = m.groupRowsBy(vals.length, (i) => vals[i]);
  assert.deepStrictEqual(g.map((x) => [x.key, x.isBlank, x.indices]), [['(빈 값)', false, [0, 2]], ['', true, [1, 3]]]);
  const used = new Set();
  assert.deepStrictEqual(g.map((x) => m.sanitizeSheetName(x.label, used)), ['(빈 값)', '(빈 값) (2)']);
});

console.log(`\n[merge-sources] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
