// tests/unit/form-engine.test.js
// 서식 엔진 단위 테스트: 실제 서식(암환자 의료비 등록 신청서 .hwp)과 그 표 구조로 만든 .hwpx로
// 기입 위치 인식 → 값 채우기 → 다시 읽어 확인. node tests/unit/form-engine.test.js

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const { readHwpTables } = require('../../hwp-reader.js');
const { parseXml, applyEdits, attr } = require('../../hwpx-xml.js');
const { modelFromHwpx, modelFromHwpTables } = require('../../form-model.js');
const { recognize, summarize } = require('../../form-recognize.js');
const { resolveValues, matchLabel, groupKind, parseRrn } = require('../../field-dictionary.js');
const { fillHwpx, planFill, previewTexts, padInto, checkInto, displayWidth } = require('../../form-fill.js');
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

const FORM = path.join(__dirname, '..', 'fixtures', 'forms', '암환자_의료비_등록_신청서.hwp');
const hwpTables = readHwpTables(fs.readFileSync(FORM));
const hwpModel = modelFromHwpTables(hwpTables);
const hwpSpots = recognize(hwpModel);
const hwpx = buildHwpx(hwpTables.map((t) => ({ cells: t.cells })), ['위와 같이 신청합니다.', '          년        월        일', '{{성명}} 귀하']);
const model = modelFromHwpx(hwpx);
const spots = recognize(model);

const find = (list, pred) => {
  const s = list.find(pred);
  if (!s) throw new Error('기입 위치를 찾지 못함');
  return s;
};
const byLabel = (list, group, label) => find(list, (s) => (s.group || '').includes(group) && s.label === label);
const cellText = (m, tableId, row, col) => {
  const t = m.tableById.get(tableId);
  const c = t.cells.find((x) => x.row === row && x.col === col);
  return c ? c.text : null;
};

/* ---------- 사전 ---------- */

test('사전: 묶음 머리글로 같은 "성명"을 환자·신청인으로 나눈다', () => {
  assert.strictEqual(groupKind('등록\n대상자'), 'patient');
  assert.strictEqual(groupKind('신청인\n(보호자 등)'), 'applicant');
  assert.strictEqual(groupKind('진단정보'), 'diagnosis');
  assert.strictEqual(matchLabel('성 명', 'patient').field, 'patient.name');
  assert.strictEqual(matchLabel('성명', 'applicant').field, 'applicant.name');
  assert.strictEqual(matchLabel('전화번호', 'diagnosis').field, 'org.phone');
  assert.strictEqual(matchLabel('가구원수\n(환자 포함)', 'patient').field, 'patient.household');
  assert.strictEqual(matchLabel('연락처', null).confidence, 'check');
  assert.strictEqual(matchLabel('세대주 성명', null), null);
});

test('사전: 주민등록번호 → 생년월일·성별, 서명란 이름은 신청 방식에 따라', () => {
  assert.deepStrictEqual(parseRrn('900101-1234567'), { y: 1990, m: 1, d: 1, sex: '남' });
  assert.deepStrictEqual(parseRrn('0503154234567'), { y: 2005, m: 3, d: 15, sex: '여' });
  assert.strictEqual(parseRrn('901301-1234567'), null);
  const today = new Date(2026, 10, 3);
  const self = resolveValues({ values: { 'patient.name': '홍길동', 'patient.rrn': '9001011234567', 'applicant.name': '김보호' }, applicantMode: 'self', today });
  assert.strictEqual(self['patient.rrn'], '900101-1234567');
  assert.strictEqual(self['patient.birth'], '1990. 1. 1.');
  assert.strictEqual(self['patient.age'], '36');
  assert.strictEqual(self['auto.signer'], '홍길동');
  assert.strictEqual(self['applicant.name'], '', '본인 신청이면 신청인 칸은 비움');
  assert.strictEqual(self['patient.holder'], '홍길동', '예금주를 비우면 환자 성명');
  const proxy = resolveValues({ values: { 'patient.name': '홍길동', 'applicant.name': '김보호', 'applicant.relation': '배우자' }, applicantMode: 'proxy', today });
  assert.strictEqual(proxy['auto.signer'], '김보호');
  assert.strictEqual(proxy['auto.signerRelation'], '배우자');
});

/* ---------- 인식 ---------- */

test('인식(실제 .hwp): 기입 위치 44곳 — 표 칸 16 · 칸 안 빈자리 10 · 선택 13 · 단위 1 · 날짜 1 · 서명 3', () => {
  const s = summarize(hwpSpots);
  assert.strictEqual(s.total, 44, JSON.stringify(s));
  assert.deepStrictEqual(s.byKind, { choice: 13, cell: 16, inline: 10, unit: 1, date: 1, sign: 3 });
  assert.strictEqual(s.check + s.none, 0, '확인 필요·비워 둠 없이 모두 정해짐');
});

test('인식: 같은 머리글을 묶음으로 구분(대상자 성명·신청인 성명·진단기관 전화번호)', () => {
  assert.strictEqual(byLabel(hwpSpots, '등록 대상자', '성명').field, 'patient.name');
  assert.strictEqual(byLabel(hwpSpots, '신청인', '성명').field, 'applicant.name');
  assert.strictEqual(byLabel(hwpSpots, '등록 대상자', '전화번호').field, 'patient.phone');
  assert.strictEqual(byLabel(hwpSpots, '신청인', '전화번호').field, 'applicant.phone');
  assert.strictEqual(byLabel(hwpSpots, '진단정보', '전화번호').field, 'org.phone');
  assert.strictEqual(byLabel(hwpSpots, '진단정보', '소재지').field, 'org.addr');
  assert.strictEqual(byLabel(hwpSpots, '진단정보', '진단기관').field, 'org.name');
});

test('인식: 칸 안 빈자리(금융기관·예금주·계좌번호), 단위 앞(가구원수 … 명), 서명 줄 3곳', () => {
  ['금융기관', '예금주', '계좌번호'].forEach((l) => assert.strictEqual(byLabel(hwpSpots, '등록 대상자', l).kind, 'inline'));
  const unit = find(hwpSpots, (s) => s.kind === 'unit');
  assert.strictEqual(unit.field, 'patient.household');
  assert.strictEqual(unit.unit, '명');
  assert.deepStrictEqual(hwpSpots.filter((s) => s.kind === 'sign').map((s) => s.label), ['신청인(대리신청인)', '동의자', '민원인']);
  assert.strictEqual(find(hwpSpots, (s) => s.label === '신청인과의 관계').field, 'auto.signerRelation');
});

test('인식: 선택 표시 — 이름에 [소아]가 든 보기, 괄호 안 하위 선택은 조건부로 따로', () => {
  const kind = find(hwpSpots, (s) => s.kind === 'choice' && s.question === '신청구분');
  assert.deepStrictEqual(kind.options.map((o) => o.text), ['건강보험가입자[소아]', '의료급여수급자[소아]', '건강보험가입자[성인]', '의료급여수급자[성인]', '폐암환자[성인]']);
  const sub = find(hwpSpots, (s) => s.kind === 'choice' && s.question === '기지원자');
  assert.deepStrictEqual(sub.options.map((o) => o.text), ['정액금', '실비']);
  assert.deepStrictEqual(sub.dependsOn, { spotId: kind.id, option: 4 }, '폐암환자[성인]을 고를 때');
  const med = find(hwpSpots, (s) => s.kind === 'choice' && s.question === '의료급여수급자 세부');
  assert.strictEqual(med.options.length, 4);
  const many = find(hwpSpots, (s) => s.kind === 'choice' && s.options.length === 8);
  assert.strictEqual(many.multiple, true);
});

test('인식: "있음 (기관:  , 금액:  )" 빈자리는 "있음"을 고를 때만, 출력할 때 묻는 칸', () => {
  const gov = find(hwpSpots, (s) => s.kind === 'choice' && s.question === '국가지원금');
  const org = find(hwpSpots, (s) => s.kind === 'inline' && s.label === '기관');
  assert.deepStrictEqual(org.dependsOn, { spotId: gov.id, option: 1 });
  assert.strictEqual(org.ask, true);
});

test('인식(.hwpx): 같은 표 구조의 .hwpx에서도 같은 기입 위치 + 본문 날짜 틀·{{변수}}', () => {
  const fromHwpx = summarize(spots.filter((s) => model.paragraphById.get(s.paragraphId).cell));
  assert.deepStrictEqual(fromHwpx.byKind, summarize(hwpSpots).byKind);
  const bodyDate = spots.filter((s) => s.kind === 'date');
  assert.strictEqual(bodyDate.length, 2);
  const ph = find(spots, (s) => s.kind === 'placeholder');
  assert.strictEqual(ph.field, 'patient.name');
});

/* ---------- 채우기 ---------- */

const today = new Date(2026, 10, 3);
const input = {
  values: {
    'patient.name': '홍길동',
    'patient.rrn': '900101-1234567',
    'patient.addr': '[예시] 서울시 중구 세종대로 1 & 2',
    'patient.phone': '02-000-0000',
    'patient.mobile': '010-0000-0000',
    'patient.email': 'test@example.com',
    'patient.household': '3',
    'patient.bank': '[예시은행]',
    'patient.account': '000-00-000000',
    'diagnosis.name': '[예시 상병]',
    'diagnosis.code': 'C00.0',
    'diagnosis.firstDate': '2026. 1. 2.',
    'applicant.name': '김보호',
    'applicant.relation': '배우자',
  },
  org: { name: '[예시 병원]', phone: '000-000-0000', addr: '[예시 소재지]' },
  applicantMode: 'proxy',
  today,
};
const values = resolveValues(input);
const choiceOf = (q) => find(spots, (s) => s.kind === 'choice' && s.question === q);
const answers = {
  values,
  choices: {
    [choiceOf('신청구분').id]: [4],
    [choiceOf('기지원자').id]: [1],
    [choiceOf('국가지원금').id]: [1],
    [choiceOf('개인‧단체 후원').id]: [0],
  },
  asks: {},
};
answers.asks[find(spots, (s) => s.kind === 'inline' && s.label === '기관' && s.dependsOn && s.dependsOn.spotId === choiceOf('국가지원금').id).id] = '[예시 기관]';
const out = fillHwpx(hwpx, spots, answers);
const filled = modelFromHwpx(out.buffer);
const tid = model.tables[0].id;

test('채우기: 표 칸에 값(대상자·신청인·진단·기관), 빈 칸 세 가지 모양 모두', () => {
  assert.strictEqual(cellText(filled, tid, 3, 2), '홍길동');
  assert.strictEqual(cellText(filled, tid, 3, 4), '900101-1234567');
  assert.strictEqual(cellText(filled, tid, 4, 2), '[예시] 서울시 중구 세종대로 1 & 2');
  assert.strictEqual(cellText(filled, tid, 9, 2), '김보호');
  assert.strictEqual(cellText(filled, tid, 10, 2), '배우자');
  assert.strictEqual(cellText(filled, tid, 12, 2), '[예시 병원]');
  assert.strictEqual(cellText(filled, tid, 13, 2), '000-000-0000');
  assert.strictEqual(cellText(filled, tid, 11, 4), 'C00.0');
});

test('채우기: 칸 안 빈자리는 빈칸 너비를 크게 바꾸지 않음, 예금주는 비우면 환자 성명, 단위 앞 "3 명"', () => {
  const t = cellText(filled, tid, 6, 2);
  assert.ok(/금융기관: \[예시은행\] +예금주: 홍길동 +계좌번호: 000-00-000000$/.test(t), t);
  const before = cellText(model, tid, 6, 2);
  assert.ok(Math.abs(displayWidth(t.split('계좌번호')[0]) - displayWidth(before.split('계좌번호')[0])) <= 2, '앞부분 너비 유지');
  assert.strictEqual(cellText(filled, tid, 7, 4).trim(), '3 명');
});

test('채우기: 고른 보기만 [ V ], 하위 선택은 부모를 고를 때만, 여러 줄 칸도', () => {
  const kind = cellText(filled, tid, 2, 1);
  assert.strictEqual((kind.match(/\[ V \]/g) || []).length, 2, kind);
  assert.ok(kind.includes('[ V ] 폐암환자[성인]'), kind);
  assert.ok(/\[ V \] 실비/.test(kind), kind);
  assert.ok(kind.includes('[   ] 정액금'), kind);
  const gov = cellText(filled, tid, 15, 2);
  assert.ok(gov.includes('[   ] 없음') && gov.includes('[ V ] 있음'), gov);
  assert.ok(/기관: \[예시 기관\]/.test(gov), '있음을 골라 기관 칸 채움');
  const priv = cellText(filled, tid, 17, 2);
  assert.ok(priv.includes('[ V ] 없음') && !/내용: \S/.test(priv), '없음을 골라 내용 칸은 비움');
});

test('채우기: 날짜 틀·서명 줄(대리 신청 → 신청인 성명)·신청인과의 관계·{{변수}}', () => {
  const sign = cellText(filled, tid, 20, 0);
  assert.ok(/2026\s*년\s+11월\s+3일/.test(sign.replace(/ +/g, ' ').replace(/(\d) 년/, '$1 년')) || /2026년\s+11월\s+3일/.test(sign), sign);
  assert.ok(/신청인\(대리신청인\)： 김보호 +\(서명 또는 인\)/.test(sign), sign);
  assert.ok(/신청인과의 관계 : 배우자 +\(대리신청의 경우\)/.test(sign), sign);
  const body = filled.paragraphs.filter((p) => !p.cell).map((p) => p.text);
  assert.ok(body.some((t) => /2026년\s+11월\s+3일/.test(t)), body.join(' | '));
  assert.ok(body.includes('홍길동 귀하'), body.join(' | '));
});

test('채우기: 본인 신청이면 신청인 칸은 비우고 서명란은 환자 성명', () => {
  const v = resolveValues({ ...input, applicantMode: 'self' });
  const res = modelFromHwpx(fillHwpx(hwpx, spots, { values: v }).buffer);
  assert.strictEqual(cellText(res, tid, 9, 2), '');
  assert.ok(/： 홍길동 +\(서명/.test(cellText(res, tid, 20, 0)));
  assert.ok(/신청인과의 관계 : +\(대리신청의 경우\)/.test(cellText(res, tid, 20, 0)), '관계 칸은 비움');
});

test('채우기: 바꾼 자리 밖의 XML은 그대로, mimetype은 압축하지 않음, 미리보기 글자 = 실제 결과', () => {
  const a = new PizZip(hwpx).file('Contents/section0.xml').asText();
  const b = new PizZip(out.buffer).file('Contents/section0.xml').asText();
  const strip = (x) => x.replace(/<hp:t>[^<]*<\/hp:t>|<hp:t\/>|<hp:t><\/hp:t>|<hp:run charPrIDRef="3"\/>|<hp:run charPrIDRef="3">/g, '');
  assert.strictEqual(strip(a).replace(/<\/hp:run>/g, ''), strip(b).replace(/<\/hp:run>/g, ''));
  const zip = new PizZip(out.buffer);
  assert.strictEqual(zip.file('mimetype').asText(), 'application/hwp+zip');
  // zip 첫 항목이 mimetype이고 압축 방식 0(저장만)인지 파일 머리 바이트로 확인
  assert.strictEqual(out.buffer.readUInt32LE(0), 0x04034b50);
  assert.strictEqual(out.buffer.readUInt16LE(8), 0, 'mimetype 압축 안 함');
  assert.strictEqual(out.buffer.slice(30, 38).toString(), 'mimetype');
  const preview = previewTexts(model, planFill(model, spots, answers));
  preview.forEach((text, pid) => assert.strictEqual(filled.paragraphById.get(pid).text, text, pid));
});

/* ---------- 도구 ---------- */

test('도구: 빈칸 맞추기·체크 표시·너비', () => {
  assert.strictEqual(checkInto(3), ' V ');
  assert.strictEqual(checkInto(2), 'V ');
  assert.strictEqual(checkInto(0), 'V');
  assert.strictEqual(displayWidth('홍길동 A1'), 9);
  assert.strictEqual(padInto(10, '홍길동'), ' 홍길동   ');
  assert.strictEqual(padInto(3, '홍길동'), ' 홍길동 ');
});

test('XML 편집기: 원문 위치 보존·문자 참조·겹침 방지', () => {
  const xml = '<a x="1 &amp; 2"><b>가&amp;나</b><c/></a>';
  const root = parseXml(xml);
  const a = root.children[0];
  assert.strictEqual(attr(a, 'x'), '1 & 2');
  const text = a.children[0].children[0];
  assert.strictEqual(xml.slice(text.start, text.end), '가&amp;나');
  assert.strictEqual(applyEdits(xml, [{ start: text.start, end: text.end, text: '다' }]), '<a x="1 &amp; 2"><b>다</b><c/></a>');
  assert.throws(() => applyEdits(xml, [{ start: 1, end: 5, text: '' }, { start: 3, end: 6, text: '' }]));
});

console.log(`\n[form-engine] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
