// tests/unit/auto-match.test.js
// 컬럼 이름 자동 연결 단위 테스트: node tests/unit/auto-match.test.js

const assert = require('assert');
const m = require('../../auto-match.js');
const { MANY_TARGET, MANY_SOURCE } = require('../e2e/fixtures.js');

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

const SYN = [
  ['이름', '성명', '성함'],
  ['전화번호', '연락처', '휴대폰번호'],
  ['email', '메일주소'],
  ['부서', '소속', '부서명'],
  ['직급', '직책', '직위'],
  ['생년월일', '생일'],
  ['이메일', '이메일주소'],
];
const matches = (res) => Object.fromEntries(Object.entries(res).map(([b, r]) => [b, r.match]));

test('타겟 순서와 무관하게 같은 이름 먼저: [성명, 이름] ← [이름]이면 이름=이름, 성명은 비움', () => {
  const r1 = m.autoMatchAll(['성명', '이름'], ['이름'], SYN);
  assert.deepStrictEqual(matches(r1), { 성명: null, 이름: '이름' });
  assert.strictEqual(r1['이름'].reason, 'exact');
  const r2 = m.autoMatchAll(['이름', '성명'], ['이름'], SYN);
  assert.deepStrictEqual(matches(r2), { 이름: '이름', 성명: null });
});

test('예전 방식(앞에서부터 하나씩)은 순서에 따라 틀렸던 경우', () => {
  // findAutoMatch를 타겟 순서대로 부르면 성명이 동의어로 "이름"을 먼저 가져간다 — autoMatchAll을 써야 하는 이유
  const used = new Set();
  const first = m.findAutoMatch('성명', ['이름'], used, SYN);
  assert.strictEqual(first.match, '이름');
});

test('컬럼이 많아도 타겟 순서를 뒤섞어 같은 결과', () => {
  const base = matches(m.autoMatchAll(MANY_TARGET, MANY_SOURCE, SYN));
  const shuffled = MANY_TARGET.slice().reverse();
  const rotated = [...MANY_TARGET.slice(17), ...MANY_TARGET.slice(0, 17)];
  assert.deepStrictEqual(matches(m.autoMatchAll(shuffled, MANY_SOURCE, SYN)), base);
  assert.deepStrictEqual(matches(m.autoMatchAll(rotated, MANY_SOURCE, SYN)), base);
  assert.strictEqual(base['성명'], '이름');
  assert.strictEqual(base['주소'], null);
});

test('정규화 이름이 같은 원본이 여럿: 글자까지 같은 것이 하나면 그것, 아니면 모호(첫 번째를 고르지 않음)', () => {
  assert.strictEqual(m.autoMatchAll(['비고'], ['비 고', '비고'], SYN)['비고'].match, '비고');
  const r = m.autoMatchAll(['Email'], ['email', 'EMAIL', '메일주소'], SYN);
  assert.strictEqual(r.Email.match, null);
  assert.deepStrictEqual(r.Email.candidates, ['email', 'EMAIL']);
  assert.ok(r.Email.note.includes('이름이 같은 원본 컬럼 2개'));
});

test('같은 원본을 두 타겟 컬럼이 같은 정도로 원하면 둘 다 연결하지 않고 이유를 남김', () => {
  const r = m.autoMatchAll(['이름', '성함'], ['성명'], SYN);
  assert.deepStrictEqual(matches(r), { 이름: null, 성함: null });
  assert.deepStrictEqual(r['이름'].candidates, ['성명']);
  assert.ok(r['이름'].note.includes('"성함"'));
});

test('동의어 후보가 여럿이면 모호, 하나면 연결', () => {
  const r = m.autoMatchAll(['연락처'], ['전화번호', '휴대폰번호'], SYN);
  assert.strictEqual(r['연락처'].match, null);
  assert.ok(r['연락처'].note.includes('동의어 후보 2개'));
  assert.strictEqual(m.autoMatchAll(['연락처'], ['전화번호', '주소'], SYN)['연락처'].reason, 'synonym');
});

test('비슷한 이름: 하나뿐이면 연결(fuzzy), 동률이면 후보로', () => {
  const r = m.autoMatchAll(['입사릴', '주소'], ['입사일', '주소1', '주소2'], SYN);
  assert.strictEqual(r['입사릴'].match, '입사일');
  assert.strictEqual(r['입사릴'].reason, 'fuzzy');
  assert.strictEqual(r['주소'].match, null);
  assert.deepStrictEqual(r['주소'].candidates, ['주소1', '주소2']);
});

test('비슷한 이름: 같은 원본을 원하는 타겟 중 더 가까운 쪽이 가져감', () => {
  // 직급수당 ← 직급수당액(거리 1), 직책수당 ← 직급수당액(거리 2): 순서와 무관하게 직급수당이 가져간다
  const r1 = m.autoMatchAll(['직책수당', '직급수당'], ['직급수당액'], []);
  const r2 = m.autoMatchAll(['직급수당', '직책수당'], ['직급수당액'], []);
  assert.deepStrictEqual(matches(r1), { 직책수당: null, 직급수당: '직급수당액' });
  assert.deepStrictEqual(matches(r2), matches(r1));
});

test('자모 단위 오타(성명↔승명)도 연결', () => {
  assert.strictEqual(m.autoMatchAll(['승명'], ['성명', '부서'], [])['승명'].match, '성명');
});

test('값 열쇠: 날짜·숫자·글자 표기 차이를 같은 값으로', () => {
  assert.strictEqual(m.valueKey(new Date(2026, 7, 1)), '2026-08-01');
  assert.strictEqual(m.valueKey('2026.8.1'), '2026-08-01');
  assert.strictEqual(m.valueKey('2026년 8월 1일'), '2026-08-01');
  assert.strictEqual(m.valueKey('100,000'), '100000');
  assert.strictEqual(m.valueKey(100000.0), '100000');
  assert.strictEqual(m.valueKey(' 상품 A '), '상품a');
  assert.strictEqual(m.valueKey(''), null);
});

const valuesOf = (b, a) => ({ b: (k) => b[k] || [], a: (k) => a[k] || [] });

test('값으로 잇기: 이름으로 못 정한 컬럼을 타겟 예시 값이 겹치는 원본 컬럼에(확인 필요)', () => {
  const A = { 고객명: ['홍길동', '김철수'], 가입일: [new Date(2026, 7, 1), new Date(2026, 7, 2)], 생년월일: [new Date(1990, 0, 2), new Date(1985, 4, 5)], 납입금액: [100000, 200000] };
  const B = { 접수일자: ['2026-08-01', '2026-08-02'], 금액: ['100,000', '200,000'], 지점: ['서울', '서울'] };
  const r = m.autoMatchAll(['접수일자', '금액', '지점', '고객명'], Object.keys(A), [], valuesOf(B, A));
  assert.strictEqual(r.고객명.reason, 'exact');
  assert.deepStrictEqual([r.접수일자.match, r.접수일자.reason, r.접수일자.valueHit], ['가입일', 'value', { hit: 2, total: 2 }]);
  assert.strictEqual(r.금액.match, '납입금액');
  assert.strictEqual(r.지점.match, null, '값이 한 종류뿐이면 잇지 않음');
  assert.ok(m.needsCheck('value') && m.needsCheck('fuzzy') && !m.needsCheck('exact'));
});

test('값으로 잇기: 값이 같은 원본 컬럼이 둘이면 잇지 않고 후보로, 이미 이름으로 이은 원본은 가져가지 않음', () => {
  const A = { 신청일: ['2026-01-01', '2026-01-02'], 접수일: ['2026-01-01', '2026-01-02'], 연락처: ['010-1', '010-2'] };
  const B = { 날짜: ['2026-01-01', '2026-01-02'], 연락처: ['010-1', '010-2'], 전화: ['010-1', '010-2'] };
  const r = m.autoMatchAll(['날짜', '연락처', '전화'], Object.keys(A), [], valuesOf(B, A));
  assert.strictEqual(r.날짜.match, null);
  assert.deepStrictEqual(r.날짜.candidates.sort(), ['신청일', '접수일']);
  assert.strictEqual(r.연락처.match, '연락처');
  assert.strictEqual(r.전화.match, null);
});

test('값으로 잇기: 일부만 겹치면 잇지 않고 후보로, 이름이 모호했던 컬럼은 그 후보 안에서 값으로 고름', () => {
  const A = { 비고: ['가', '나', '다', '라'], 메모: ['x', 'y'], 비고_1: ['p', 'q'] };
  const partial = m.autoMatchAll(['특이사항'], ['비고'], [], valuesOf({ 특이사항: ['가', '마', '바'] }, A));
  assert.strictEqual(partial.특이사항.match, null);
  assert.deepStrictEqual(partial.특이사항.candidates, ['비고']);
  const tie = m.autoMatchAll(['비고'], ['비고', '비고'], [], null);
  assert.strictEqual(tie.비고.match, null);
  const syn = [['비고', '메모', '참고']];
  const pick = m.autoMatchAll(['참고'], ['비고', '메모'], syn, valuesOf({ 참고: ['x', 'y'] }, A));
  assert.deepStrictEqual([pick.참고.match, pick.참고.reason], ['메모', 'value']);
});

test('초성', () => {
  assert.strictEqual(m.getChoseong('연락처 A'), 'ㅇㄹㅊ A');
});

console.log(`\n[auto-match] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
