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

test('초성', () => {
  assert.strictEqual(m.getChoseong('연락처 A'), 'ㅇㄹㅊ A');
});

console.log(`\n[auto-match] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
