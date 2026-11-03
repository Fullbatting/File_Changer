// tests/unit/common-forms.test.js
// 정부 공통서식 모양 단위 테스트: 보건복지부 긴급복지 서식(수급계좌 입금 신청서·금융정보 제공 동의서·
// 진료비 확인 위임장·요양급여 대상여부 확인 요청서)의 표 구조를 본떠 만든 .hwpx로 인식·채우기를 확인한다.
// 글자는 모두 가짜 예시다. node tests/unit/common-forms.test.js

const assert = require('assert');
const { modelFromHwpx } = require('../../form-model.js');
const { recognize } = require('../../form-recognize.js');
const { resolveValues, groupKind, matchLabel } = require('../../field-dictionary.js');
const { fillHwpx } = require('../../form-fill.js');
const { buildHwpx } = require('./hwpx-fixture.js');

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

const c = (row, col, text, colSpan = 1, rowSpan = 1) => ({ row, col, text, colSpan, rowSpan });

// [서식 5호] 같은 모양: 머리글과 값을 한 칸에 쓰는 넓은 칸, 주소 칸 안 "(전화:   )", 시청 칸, 처리 절차 그림
const DEPOSIT = {
  cells: [
    c(0, 0, '접수번호', 2), c(0, 2, '접수일자'), c(0, 3, ''), c(0, 4, '처리기간'), c(0, 5, '즉시'),
    c(1, 0, '신청인\n(지원대상자)', 2, 2), c(1, 2, '성  명', 2), c(1, 4, '생년월일', 2),
    c(2, 2, '주  소\n(전화:                    )', 4),
    c(3, 0, '긴급지원\n수급계좌', 2, 2), c(3, 2, '금융기관', 2), c(3, 4, '예금주', 2),
    c(4, 2, '계좌번호', 4),
    c(5, 0, '                     20    년      월      일\n            신청인:                  (서명 또는 인)', 6),
    c(6, 0, ' 시장ㆍ군수ㆍ구청장 귀하', 6),
    c(7, 0, '신청서 작성', 2), c(7, 2, '➡'), c(7, 3, '접  수'), c(7, 4, '➡'), c(7, 5, '계좌번호 등록'),
    c(8, 0, '신청인', 2), c(8, 2, ''), c(8, 3, '시ㆍ군ㆍ구'), c(8, 4, ''), c(8, 5, ''),
  ],
};

// [별지 4] 같은 모양: 머리글 아래 빈 줄, "ㅇ 자택 :  ㅇ 휴대전화 :", 위임받는 쪽(시청)
const PROXY = {
  cells: [
    c(0, 0, '진료비 확인 요청 위임장(보장기관)', 3),
    c(1, 0, '위임하는 사람\n(지원대상자)', 1, 3), c(1, 1, '성명\n'), c(1, 2, '연락처\nㅇ 자택 :\nㅇ 휴대전화 : '),
    c(2, 1, '생년월일\n', 2),
    c(3, 1, '주소\n', 2),
    c(4, 0, '위임받는 사람\n(수임인)', 1, 2), c(4, 1, '보장기관명칭\n'), c(4, 2, '보장기관의 장\n'),
    c(5, 1, '주소\n'), c(5, 2, '연락처'),
    c(6, 0, '※ [  ]에는 해당되는 곳에 √ 표시를 합니다.', 3),
  ],
};

// [서식 3호] 같은 모양: 머리글 줄 + 빈 줄, 주민등록번호 칸에는 "-"가 미리 찍혀 있음, 긴 서명 머리글.
// 점선 상자가 떠 있는 빈 칸은 쓸 수 있고, 표가 들어앉은 칸은 쓰지 않는다.
const CONSENT = {
  cells: [
    c(0, 0, '관 계'), c(0, 1, '성 명'), c(0, 2, '주민등록번호\n(외국인등록번호 등)'), c(0, 3, '금융정보 등의\n제공을 동의함1),2)\n(한글정자 서명\n또는 무인･인감)'),
    c(1, 0, ''), c(1, 1, '{RECT}'), c(1, 2, '- '), c(1, 3, ''),
    c(2, 0, '비고'), c(2, 1, '{TBL}', 3),
  ],
};

// [별지 3] 같은 모양: 한 줄짜리 묶음 머리글, 통보받을 자(신청인 쪽), 서명 칸이 따로 있는 줄, 요청인 생년월일
const REQUEST = {
  cells: [
    c(0, 0, '긴급의료지원 대상자', 2), c(0, 2, '성명', 2), c(0, 4, '주민(외국인)등록번호', 2),
    c(1, 0, '', 6),
    c(2, 0, '진료비 확인 결과\n통보받을 자', 2, 2), c(2, 2, '성명', 2), c(2, 4, '긴급지원대상자와의 관계', 2),
    c(3, 2, '연락처', 2), c(3, 4, '주소', 2),
    c(4, 0, '요청인\n(법정대리인 또는 가족)', 2), c(4, 2, '생년월일', 2), c(4, 4, '진료받은 사람과의 관계', 2),
    c(5, 0, '요청인', 4), c(5, 4, '(서명 또는 인)    ', 2),
  ],
};

// 표 번호: t1 입금 신청서 · t2 위임장 · t3 동의서(t4는 그 안에 들어앉은 표) · t5 확인 요청서
const buffer = buildHwpx([DEPOSIT, PROXY, CONSENT, REQUEST]);
const model = modelFromHwpx(buffer);
const spots = recognize(model);
const inTable = (id) => spots.filter((s) => model.paragraphById.get(s.paragraphId).cell && model.paragraphById.get(s.paragraphId).cell.tableId === id);
const fieldsOf = (id) => inTable(id).map((s) => `${s.kind}:${s.field || '-'}`);

const values = resolveValues({
  values: {
    'patient.name': '홍길동',
    'patient.rrn': '900101-1234567',
    'patient.addr': '[예시] ○○시 ○○로 12',
    'patient.phone': '031-000-0000',
    'patient.mobile': '010-1111-2222',
    'patient.bank': '[예시은행]',
    'patient.account': '000-00-000000',
    'applicant.name': '김보호',
    'applicant.rrn': '600101-2345678',
    'applicant.relation': '배우자',
    'applicant.phone': '010-3333-4444',
    'applicant.addr': '[예시] 같은 주소',
  },
  applicantMode: 'proxy',
  today: new Date(2026, 10, 3),
});
const filled = modelFromHwpx(fillHwpx(buffer, spots, { values, choices: {} }).buffer);
const cellText = (tableId, row, col) => filled.tableById.get(tableId).cells.find((x) => x.row === row && x.col === col).text;

test('사전: "신청인(지원대상자)"·"위임하는 사람(지원대상자)"는 환자, 요청인·신고인·통보받을 자는 신청인, 보장기관·직장은 따로', () => {
  assert.strictEqual(groupKind('신청인\n(지원대상자)'), 'patient');
  assert.strictEqual(groupKind('위임하는 사람\n(지원대상자)'), 'patient');
  assert.strictEqual(groupKind('진료받은 사람'), 'patient');
  assert.strictEqual(groupKind('요청인 또는 신고인'), 'applicant');
  assert.strictEqual(groupKind('요청인\n(법정대리인 또는 가족)'), 'applicant');
  assert.strictEqual(groupKind('진료비 확인 결과 통보받을 자'), 'applicant');
  assert.strictEqual(groupKind('진료비 확인 요청인\n(보장기관)'), 'office');
  assert.strictEqual(groupKind('위임받는 사람\n(수임인)'), 'office');
  assert.strictEqual(groupKind('취업상태'), 'work');
  assert.strictEqual(matchLabel('ㅇ 휴대전화', 'patient').field, 'patient.mobile');
  assert.strictEqual(matchLabel('주민(외국인)등록번호', 'patient').field, 'patient.rrn');
  assert.strictEqual(matchLabel('생년월일', 'applicant').field, 'applicant.birth');
  assert.strictEqual(matchLabel('성명', 'office'), null);
  assert.strictEqual(values['applicant.birth'], '1960. 1. 1.');
});

test('인식(입금 신청서 모양): 넓은 칸 머리글 뒤에 이어 쓰기 6곳 + 주소 칸 안 전화 + 날짜 + 서명, 접수·귀하·처리 절차는 빼기', () => {
  assert.deepStrictEqual(fieldsOf('t1'), [
    'incell:patient.name', 'incell:patient.birth', 'incell:patient.addr', 'inline:patient.phone',
    'incell:patient.bank', 'incell:patient.holder', 'incell:patient.account', 'date:auto.today', 'sign:auto.signer',
  ]);
  assert.ok(inTable('t1').filter((s) => s.kind === 'incell').every((s) => s.append));
});

test('채우기(입금 신청서 모양): "성  명  홍길동", 전화 빈자리, "20    년" 틀은 2026년으로', () => {
  assert.strictEqual(cellText('t1', 1, 2), '성  명  홍길동');
  assert.strictEqual(cellText('t1', 1, 4), '생년월일  1990. 1. 1.');
  assert.ok(cellText('t1', 2, 2).startsWith('주  소  [예시] ○○시 ○○로 12\n(전화: 031-000-0000'), cellText('t1', 2, 2));
  assert.strictEqual(cellText('t1', 3, 4), '예금주  홍길동');
  assert.strictEqual(cellText('t1', 4, 2), '계좌번호  000-00-000000');
  const sign = cellText('t1', 5, 0);
  assert.ok(/ 2026년\s+11월\s+3일/.test(sign) && !sign.includes('202026'), sign);
  assert.ok(/신청인: 김보호\s+\(서명 또는 인\)/.test(sign), sign);
  assert.strictEqual(cellText('t1', 0, 3), '', '접수일자 칸은 시청이 쓴다');
});

test('인식·채우기(위임장 모양): 머리글 아래 빈 줄에, 자택·휴대전화 빈자리, 수임인(시청) 칸은 비워 두고 고정값 안내, 안내 문구 [  ]는 선택지 아님', () => {
  assert.deepStrictEqual(fieldsOf('t2'), ['incell:patient.name', 'inline:patient.phone', 'inline:patient.mobile', 'incell:patient.birth', 'incell:patient.addr', 'incell:-', 'incell:-', 'incell:-']);
  assert.ok(inTable('t2').filter((s) => !s.field).every((s) => s.why.includes('고정값')));
  assert.ok(!spots.some((s) => s.kind === 'choice'));
  assert.strictEqual(cellText('t2', 1, 1), '성명\n홍길동');
  assert.strictEqual(cellText('t2', 1, 2), '연락처\nㅇ 자택 : 031-000-0000\nㅇ 휴대전화 : 010-1111-2222 ');
  assert.strictEqual(cellText('t2', 4, 1), '보장기관명칭\n');
});

test('인식·채우기(동의서 모양): 머리글 줄 아래 빈 줄, "-"만 찍힌 주민등록번호 칸, 점선 상자 칸은 쓰고 표가 든 칸은 안 씀', () => {
  const t3 = inTable('t3');
  assert.deepStrictEqual(t3.map((s) => `${s.kind}:${s.label}`), ['below:관 계', 'below:성 명', 'below:주민등록번호 (외국인등록번호 등)']);
  assert.ok(t3.every((s) => !s.group), '첫 칸 "관 계"는 묶음 머리글이 아니다');
  assert.strictEqual(cellText('t3', 1, 1), '홍길동');
  assert.strictEqual(cellText('t3', 1, 2), '900101-1234567');
  assert.strictEqual(cellText('t3', 1, 3), '', '서명 칸은 손으로');
  assert.ok(!spots.some((s) => model.paragraphById.get(s.paragraphId).hasObject), '표가 들어앉은 문단에는 쓰지 않는다');
});

test('인식·채우기(확인 요청서 모양): 한 줄 묶음 머리글, 통보받을 자 = 신청인, 요청인 생년월일, 따로 있는 서명 칸', () => {
  assert.deepStrictEqual(fieldsOf('t5'), [
    'incell:patient.name', 'incell:patient.rrn', 'incell:applicant.name', 'incell:applicant.relation', 'incell:applicant.phone', 'incell:applicant.addr',
    'incell:applicant.birth', 'incell:applicant.relation', 'sign:auto.signer',
  ]);
  assert.strictEqual(cellText('t5', 0, 2), '성명  홍길동');
  assert.strictEqual(cellText('t5', 0, 4), '주민(외국인)등록번호  900101-1234567');
  assert.strictEqual(cellText('t5', 4, 2), '생년월일  1960. 1. 1.');
  assert.strictEqual(cellText('t5', 5, 4), '김보호  (서명 또는 인)    ');
});

console.log(`\n[common-forms] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
