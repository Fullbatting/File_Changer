// tests/unit/autofill.test.js
// 서식 자동 기입: 메일머지 필드 사전(merge-fields.js), .hwp 필드·목차 읽기(hwp-reader.js), 서식 보관함(form-library.js),
// 채우기·인쇄(print-job.js), 서류 출력 규칙(print-plan.js), 설정(settings.js). 값은 모두 가짜 예시다. node tests/unit/autofill.test.js

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const PizZip = require('pizzip');
const mf = require('../../merge-fields.js');
const { readHwpFields } = require('../../hwp-reader.js');
const lib = require('../../form-library.js');
const job = require('../../print-job.js');
const plan = require('../../print-plan.js');
const { loadSettings, saveSettings } = require('../../settings.js');
const { buildHwp } = require('./hwp-fixture.js');
const { buildHwpx } = require('./hwpx-fixture.js');

let passed = 0;
let failed = 0;
const queue = [];
const test = (title, fn) => queue.push({ title, fn });

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'autofill-unit-'));
const NS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"';
let fid = 0;
const F = (n) => `<hp:run charPrIDRef="0"><hp:ctrl><hp:fieldBegin id="${++fid}" type="MAILMERGE" name="${n}"/></hp:ctrl><hp:t>{{${n}}}</hp:t><hp:ctrl><hp:fieldEnd beginIDRef="${fid}"/></hp:ctrl></hp:run>`;
const T = (t) => `<hp:run charPrIDRef="0"><hp:t>${t}</hp:t></hp:run>`;
const P = (...r) => `<hp:p id="0">${r.join('')}</hp:p>`;
function hwpxWith(paragraphs) {
  const zip = new PizZip();
  zip.file('mimetype', 'application/hwp+zip');
  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?><hs:sec ${NS}>${paragraphs.join('')}</hs:sec>`);
  return zip.generate({ type: 'nodebuffer' });
}
const textOf = (buffer) =>
  new PizZip(buffer)
    .file('Contents/section0.xml')
    .asText()
    .replace(/<hp:ctrl>.*?<\/hp:ctrl>/g, '')
    .replace(/<\/hp:p>/g, '\n')
    .replace(/<[^>]+>/g, '');

const INPUT = { name: '홍길동', rrn: '9001011234567', addr: '[예시 주소]', phone: '010-0000-0000', writtenDate: '2027-08-03', careStart: '2026-01-02', disease: '[예시 상병]', diseaseCode: 'C00.0', diagnosisDate: '26.1.15', staffName: '[담당자]', staffTitle: '[직위]', staffPhone: '0000' };

/* ---------- 필드 사전 ---------- */

test('필드 값: 지금까지 쓰던 메일머지 자료 모양(생년월일 6자리·주민번호 한 자리씩·작성일 두 자리씩·진료기간·진단일)', () => {
  const names = ['이름', '생년월일', '주민등록번호', '1', '7', '13', '월1', '월2', '일1', '일2', '진료기간', '진단일', '주질병', '질병분류기호', '담당자', '직위', '담당자 번호', '모르는칸'];
  const v = mf.fieldValues(names, INPUT, { extra: { 모르는칸: '[x]' } });
  assert.deepStrictEqual(
    names.map((n) => v[n]),
    ['홍길동', '900101', '900101-1234567', '9', '1', '7', '0', '8', '0', '3', '26.01.02~', '26.01.15', '[예시 상병]', 'C00.0', '[담당자]', '[직위]', '0000', '[x]']
  );
  assert.strictEqual(mf.fieldValues(['진료기간'], { careStart: '2026-01-02', careEnd: '2026-08-01' })['진료기간'], '26.01.02~26.08.01');
  const today = mf.fieldValues(['월1', '월2', '일1', '일2'], {}, { today: new Date(2026, 10, 4) });
  assert.strictEqual(['월1', '월2', '일1', '일2'].map((n) => today[n]).join(''), '1104', '작성일을 비우면 오늘');
});

test('필드 사전: 다른 이름(성명·전화번호·상병명)도, 서식에 든 필드로 물어볼 입력만', () => {
  assert.strictEqual(mf.fieldOf('성 명').id, 'name');
  assert.strictEqual(mf.fieldOf('전화번호').id, 'phone');
  assert.strictEqual(mf.fieldOf('상병명').id, 'disease');
  assert.strictEqual(mf.fieldOf('담당자번호').id, 'staffPhone');
  assert.strictEqual(mf.fieldOf('보호자'), null);
  assert.deepStrictEqual(mf.inputsFor(['이름', '4', '월1', '보호자']), { keys: ['name', 'rrn', 'writtenDate'], unknown: ['보호자'] });
  assert.deepStrictEqual(mf.inputsFor(['진료기간']).keys, ['careStart', 'careEnd']);
  assert.deepStrictEqual(mf.missingInputs(['name', 'writtenDate', 'careStart', 'careEnd'], { name: 'x' }), ['careStart'], '작성일·진료 종료일은 비어도 된다');
});

test('주민등록번호 확인·안내, 날짜 읽기', () => {
  assert.strictEqual(mf.checkRrn('900101-1234567'), '');
  assert.ok(mf.checkRrn('9001011').includes('13자리'));
  assert.ok(mf.checkRrn('901301-1234567').includes('생년월일'));
  assert.strictEqual(mf.rrnSummary('900101-1234567', new Date(2026, 10, 4)), '생년월일 900101 · 남 · 만 36세');
  assert.strictEqual(mf.rrnSummary('050301-4234567', new Date(2026, 10, 4)), '생년월일 050301 · 여 · 만 21세');
  assert.deepStrictEqual(mf.parseDate('2026.7.3'), { y: 2026, m: 7, d: 3 });
  assert.deepStrictEqual(mf.parseDate('26.07.03.'), { y: 2026, m: 7, d: 3 });
  assert.strictEqual(mf.shortDate('아무 글'), '아무 글');
});

/* ---------- .hwp 읽기 ---------- */

test('.hwp: 메일머지 필드(이름·개수, 처음 나온 순서)와 서류 목차(띄어 쓴 제목은 붙임)', () => {
  const buf = buildHwp([
    ['[서식 1호] '],
    ['현 장 확 인 서'],
    ['성명 ', { field: '이름', extra: 'GIVEN_NAME' }],
    ['2026년 ', { field: '월1' }, { field: '월2' }, '월'],
    ['[서식 2-1호] 공통서식[별지 제1호의2서식]'],
    ['소득･재산 신고서 [□신규 □변경]'],
    [{ field: '이름' }, ' ', { field: '4' }],
  ]);
  const r = readHwpFields(buf);
  assert.deepStrictEqual(r.fields, [{ name: '이름', count: 2 }, { name: '월1', count: 1 }, { name: '월2', count: 1 }, { name: '4', count: 1 }]);
  assert.deepStrictEqual(r.documents, [{ mark: '[서식 1호]', title: '현장확인서' }, { mark: '[서식 2-1호]', title: '소득･재산 신고서 [□신규 □변경]' }]);
  const plain = readHwpFields(buildHwp([['서울형 긴급복지 지원 신청서'], [{ field: '이름' }], ['개인정보 수집 동의서']], { compressed: false }));
  assert.deepStrictEqual(plain.documents.map((d) => d.title), ['서울형 긴급복지 지원 신청서', '개인정보 수집 동의서'], '머리표가 없으면 제목만으로');
});

/* ---------- 서식 보관함 ---------- */

// 지금까지 쓰던 메일머지 자료의 파일 이름(서식 내용은 쓰지 않는다)
const REAL_NAMES = [
  ['1. (세대주)의료비지원 신청서-메일머지/★의정부시청_긴급복지지원 신청서(세대주).hwp', '의정부시청', '긴급복지(국가형)', '세대주', true, '긴급복지지원 신청서'],
  ['1. (세대주)의료비지원 신청서-메일머지/★포천시청-긴급지원(의료비) 신청서(세대주).hwp', '포천시청', '긴급복지(국가형)', '세대주', true, '긴급지원(의료비) 신청서'],
  ['1. (세대주)의료비지원 신청서-메일머지/2025 긴급복지지원 신청서(세대주).hwp', '공통 서식', '긴급복지(국가형)', '세대주', false, '2025 긴급복지지원 신청서'],
  ['1. (세대주)의료비지원 신청서-메일머지/광주 광산구청-2025 긴급복지지원 신청서(세대주).hwp', '광주 광산구청', '긴급복지(국가형)', '세대주', false, '2025 긴급복지지원 신청서'],
  ['2. (세대원)의료비지원 신청서-메일머지/동두천시청-경기도형 긴급복지 신청서(세대원).hwp', '동두천시청', '경기도형 긴급복지', '세대원', false, '경기도형 긴급복지 신청서'],
  ['2. (세대원)의료비지원 신청서-메일머지/포천시청-경기도형 긴급복지(생계,의료) 신청서(세대원).hwp', '포천시청', '경기도형 긴급복지', '세대원', false, '경기도형 긴급복지(생계,의료) 신청서'],
  ['2. (세대원)의료비지원 신청서-메일머지/연천군청-2025 긴급복지지원 신청서(세대원).hwp', '연천군청', '긴급복지(국가형)', '세대원', false, '2025 긴급복지지원 신청서'],
  ['2026년 서울형긴급복지지원.hwp', '서울특별시', '서울형 긴급복지', '공통', false, '2026년 서울형긴급복지지원'],
  ['광주 광산구청-광주형 긴급지원(노랑호루라기) 서식-메일머지.hwp', '광주 광산구청', '광주형 긴급지원', '공통', false, '광주형 긴급지원(노랑호루라기) 서식'],
  ['보건소 암지원 신청서/공통_암환자의료비등록신청서(메일머지).hwp', '공통 서식', '암환자 의료비', '공통', false, '암환자의료비등록신청서'],
  ['보건소 암지원 신청서/동두천 보건소_암환자의료비등록신청서+보안각서(메일머지).hwp', '동두천 보건소', '암환자 의료비', '공통', false, '암환자의료비등록신청서+보안각서'],
];

test('파일 이름으로 분류: 제출 기관 · 지원 종류 · 세대주/세대원 · ★ · 서식 이름', () => {
  REAL_NAMES.forEach(([file, agency, kind, applicant, favorite, title]) => {
    assert.deepStrictEqual(lib.classifyFileName(file), { agency, kind, applicant, favorite, title }, file);
  });
  assert.strictEqual(lib.classifyFileName('C:\\서식\\(세대원)\\양주시청-긴급복지지원 신청서.hwpx').applicant, '세대원', '폴더 이름의 구분도 본다');
});

test('주소로 제출 기관 추천: 시·군·구가 맞는 기관, 서울 주소는 서울형도, 없는 지역은 빈 목록', () => {
  const agencies = ['공통 서식', '의정부시청', '포천시청', '광주 광산구청', '노원구청', '동두천시청', '양주시청', '연천군청', '서울특별시', '동두천 보건소'];
  assert.deepStrictEqual(lib.suggestAgencies('경기도 양주시 백석읍 [예시로] 1', agencies), ['양주시청']);
  assert.deepStrictEqual(lib.suggestAgencies('서울특별시 노원구 [예시로] 1', agencies), ['노원구청', '서울특별시']);
  assert.deepStrictEqual(lib.suggestAgencies('경기도 동두천시 [예시로]', agencies), ['동두천시청', '동두천 보건소']);
  assert.deepStrictEqual(lib.suggestAgencies('광주광역시 광산구 [예시로]', agencies), ['광주 광산구청']);
  assert.deepStrictEqual(lib.suggestAgencies('경기도 광주시 오포읍', agencies), [], '경기도 광주시는 광주 광산구청이 아니다');
  assert.deepStrictEqual(lib.suggestAgencies('서울시 강북구 [예시로]', agencies), ['서울특별시']);
  assert.deepStrictEqual(lib.suggestAgencies('', agencies), []);
});

test('보관함: 등록·변환본·고치기·기관 목록·상태·지우기, 망가진 목록 파일은 남겨 두고 새로', () => {
  const ldir = fs.mkdtempSync(path.join(dir, 'lib-'));
  const src = path.join(ldir, '양주시청-긴급복지지원 신청서(세대주).hwp');
  fs.writeFileSync(src, buildHwp([['[서식 1호]'], ['현 장 확 인 서'], [{ field: '이름' }, { field: '보호자' }]]));
  const L = lib.openLibrary(path.join(ldir, 'forms'));
  const info = lib.inspectFormFile(src);
  const e = L.add(src, { ...lib.classifyFileName(src), ...info });
  assert.deepStrictEqual(e.mapping, { 이름: { type: 'auto' }, 보호자: { type: 'ask' } });
  assert.deepStrictEqual(lib.entryStatus(e), { ready: false, needsConvert: true, label: '확인 필요', problems: ['사전에 없는 필드 1개 확인 필요(보호자)'] });
  const conv = path.join(ldir, 'conv.hwpx');
  fs.writeFileSync(conv, hwpxWith([P(F('이름'))]));
  L.setConverted(e.id, conv);
  L.update(e.id, { reviewed: true, favorite: true });
  const again = lib.openLibrary(path.join(ldir, 'forms'));
  const e2 = again.get(e.id);
  assert.ok(e2.favorite && e2.reviewed && fs.existsSync(again.hwpxPath(e2)) && fs.existsSync(again.sourcePath(e2)));
  assert.strictEqual(lib.entryStatus(e2).label, '사용 중');
  L.add(src, { agency: '공통 서식', kind: '기타', applicant: '공통', title: 'x', fields: [], documents: [] });
  assert.deepStrictEqual(again.agencies().map((a) => a.name), ['양주시청']);
  assert.deepStrictEqual(lib.openLibrary(path.join(ldir, 'forms')).agencies().map((a) => a.name), ['양주시청', '공통 서식'], '공통 서식은 맨 뒤');
  again.remove(e.id);
  assert.ok(!fs.existsSync(path.join(ldir, 'forms', 'files', e2.hwpx)));
  fs.writeFileSync(path.join(ldir, 'forms', 'index.json'), '{broken');
  const broken = lib.openLibrary(path.join(ldir, 'forms'));
  assert.ok(broken.loadError && broken.list().length === 0);
  assert.ok(fs.readdirSync(path.join(ldir, 'forms')).some((f) => f.startsWith('index.json.broken-')));
});

test('보관함 백업: 내보내기 → 다른 PC에서 가져오기, 경로를 벗어나는 이름은 무시', () => {
  const a = lib.openLibrary(fs.mkdtempSync(path.join(dir, 'a-')));
  const src = path.join(dir, '노원구청-긴급복지지원 신청서(세대원).hwpx');
  fs.writeFileSync(src, hwpxWith([P(T('성명 '), F('이름'))]));
  a.add(src, { ...lib.classifyFileName(src), ...lib.inspectFormFile(src) });
  const zipBuf = a.exportBackup();
  const b = lib.openLibrary(fs.mkdtempSync(path.join(dir, 'b-')));
  assert.deepStrictEqual(b.importBackup(zipBuf), { added: 1, replaced: 0 });
  assert.deepStrictEqual(b.importBackup(zipBuf), { added: 0, replaced: 1 });
  assert.strictEqual(b.list()[0].agency, '노원구청');
  const evil = new PizZip(zipBuf);
  const data = JSON.parse(evil.file('index.json').asText());
  data.entries[0].id = '../../evil';
  evil.file('index.json', JSON.stringify(data));
  assert.deepStrictEqual(b.importBackup(evil.generate({ type: 'nodebuffer' })), { added: 0, replaced: 0 });
  assert.throws(() => b.importBackup(Buffer.from('x')), /백업 파일이 아닙니다/);
});

test('등록 확인 줄: 분류·필드·사전에 없는 필드, 이미 등록한 서식과 읽을 수 없는 파일은 처음에 고르지 않음', () => {
  const root = fs.mkdtempSync(path.join(dir, 'import-'));
  fs.mkdirSync(path.join(root, '2. (세대원)'));
  fs.writeFileSync(path.join(root, '2. (세대원)', '양주시청-긴급복지지원 신청서.hwp'), buildHwp([[{ field: '이름' }, { field: '보호자' }]]));
  fs.writeFileSync(path.join(root, '노원구청-긴급복지지원 신청서(세대주).hwpx'), 'not a zip');
  const rows = lib.prepareImport([root], [{ sourceName: '양주시청-긴급복지지원 신청서.hwp', agency: '양주시청', applicant: '세대원' }]);
  const yang = rows.find((r) => r.agency === '양주시청');
  const nowon = rows.find((r) => r.agency === '노원구청');
  assert.deepStrictEqual([yang.applicant, yang.fields.length, yang.unknown, yang.dup, yang.include], ['세대원', 2, ['보호자'], true, false]);
  assert.ok(nowon.error && !nowon.include && nowon.fields.length === 0);
  assert.deepStrictEqual(lib.prepareImport([path.join(root, 'nothing')].filter(fs.existsSync)), []);
});

test('보관함 백업: 다른 서식의 파일 이름을 가리키는 항목은 건너뛰고, 모르는 칸은 버린다', () => {
  const a = lib.openLibrary(fs.mkdtempSync(path.join(dir, 'c-')));
  const src = path.join(dir, '연천군청-긴급복지지원 신청서(세대주).hwpx');
  fs.writeFileSync(src, hwpxWith([P(F('이름'))]));
  const e = a.add(src, { ...lib.classifyFileName(src), ...lib.inspectFormFile(src) });
  const zip = new PizZip(a.exportBackup());
  const data = JSON.parse(zip.file('index.json').asText());
  data.entries.push({ ...data.entries[0], id: 'aaaaaaaaaaaa' }); // 파일은 e의 것을 가리킴
  data.entries[0].extra = '<script>';
  data.entries[0].kind = '없는 종류';
  zip.file('index.json', JSON.stringify(data));
  const b = lib.openLibrary(fs.mkdtempSync(path.join(dir, 'd-')));
  assert.deepStrictEqual(b.importBackup(zip.generate({ type: 'nodebuffer' })), { added: 1, replaced: 0 });
  const got = b.get(e.id);
  assert.ok(!('extra' in got) && got.kind === '기타' && got.agency === '연천군청');
});

test('폴더 안의 서식 찾기: 하위 폴더까지, 한글 임시 파일(~$)·다른 파일은 빼고', () => {
  const root = fs.mkdtempSync(path.join(dir, 'tree-'));
  fs.mkdirSync(path.join(root, '1. (세대주)'));
  ['a.hwp', '1. (세대주)/b.hwpx', '1. (세대주)/~$b.hwpx', 'c.xlsx', '.hidden.hwp'].forEach((f) => fs.writeFileSync(path.join(root, f), 'x'));
  assert.deepStrictEqual(lib.findFormFiles(root).map((f) => path.relative(root, f)).sort(), [path.join('1. (세대주)', 'b.hwpx'), 'a.hwp'].sort());
});

/* ---------- 채우기 · 인쇄 ---------- */

test('채우기(메일머지 필드): 값·연도 맞춤·고정 글자·비워 둠·출력 때 입력, 필드 표시는 그대로', () => {
  const buf = hwpxWith([
    P(T('성명 '), F('이름'), T(' 생년월일 '), F('생년월일')),
    P(T('2026년 '), F('월1'), F('월2'), T('월 '), F('일1'), F('일2'), T('일')),
    P(F('4'), F('5'), F('6'), F('1'), F('2'), F('3'), F('7'), F('13')),
    P(T('의료기관 '), F('기관'), T(' / '), F('메모'), T(' / '), F('보호자')),
  ]);
  const entry = { mapping: { 기관: { type: 'fixed', text: '[예시 병원]' }, 메모: { type: 'blank' }, 보호자: { type: 'ask' } } };
  const r = job.fillForm(entry, buf, INPUT, { asks: { 보호자: '[예시 보호자]' } });
  assert.strictEqual(r.mode, 'fields');
  assert.strictEqual(textOf(r.buffer), '성명 홍길동 생년월일 900101\n2027년 08월 03일\n10190017\n의료기관 [예시 병원] /  / [예시 보호자]\n');
  const xml = new PizZip(r.buffer).file('Contents/section0.xml').asText();
  assert.strictEqual((xml.match(/fieldBegin/g) || []).length, 17);
});

test('채우기(메일머지 필드): 글상자 안·글상자 속 글상자·표 칸 안 글상자의 필드도 한 번씩만 고친다(겹치는 편집 없음)', () => {
  // 한글은 글상자를 <hp:rect> > <hp:drawText> > <hp:subList> > <hp:p>로, 바깥 문단의 런 안에 둔다
  const box = (...ps) => `<hp:run charPrIDRef="0"><hp:rect id="9"><hp:drawText><hp:subList>${ps.join('')}</hp:subList></hp:drawText></hp:rect></hp:run>`;
  const cell = (...ps) => `<hp:run charPrIDRef="0"><hp:tbl rowCnt="1" colCnt="1"><hp:tr><hp:tc><hp:subList>${ps.join('')}</hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/></hp:tc></hp:tr></hp:tbl></hp:run>`;
  const buf = hwpxWith([
    P(T('신청인 '), box(P(T('성명 '), F('이름'))), T(' 끝')),
    P(box(P(T('연락처 '), box(P(F('연락처')))))),
    P(cell(P(T('주소 '), box(P(F('주소')), P(T('2026년 '), F('월1'), F('월2'), T('월')))))),
  ]);
  const r = job.fillForm({ mapping: {} }, buf, INPUT);
  assert.strictEqual(r.mode, 'fields');
  assert.strictEqual(r.filled, 5);
  const t = textOf(r.buffer);
  assert.strictEqual(t, '신청인 성명 홍길동\n 끝\n연락처 010-0000-0000\n\n\n주소 [예시 주소]\n2027년 08월\n\n\n');
  assert.strictEqual(t.split('홍길동').length, 2);
});

test('채우기(필드 없는 서식): 자동 인식으로 표 칸을 채운다', () => {
  const buf = buildHwpx([{ cells: [{ row: 0, col: 0, text: '성명' }, { row: 0, col: 1, text: '' }, { row: 0, col: 2, text: '주민등록번호' }, { row: 0, col: 3, text: '' }] }]);
  const r = job.fillForm({ mapping: {} }, buf, INPUT);
  assert.strictEqual(r.mode, 'auto');
  const t = textOf(r.buffer);
  assert.ok(t.includes('홍길동') && t.includes('900101-1234567'), t);
});

test('인쇄: .hwp는 먼저 한글로 변환해 보관, 채운 임시 파일은 성공·실패 모두 지움, 한 서식 실패해도 계속', async () => {
  const ldir = fs.mkdtempSync(path.join(dir, 'print-'));
  const L = lib.openLibrary(path.join(ldir, 'forms'));
  const hwp = path.join(ldir, '양주시청-긴급복지지원 신청서(세대주).hwp');
  fs.writeFileSync(hwp, buildHwp([[{ field: '이름' }]]));
  const hwpx = path.join(ldir, '노원구청-긴급복지지원 신청서(세대주).hwpx');
  fs.writeFileSync(hwpx, hwpxWith([P(T('성명 '), F('이름'))]));
  const a = L.add(hwp, { ...lib.classifyFileName(hwp), ...lib.inspectFormFile(hwp) });
  const b = L.add(hwpx, { ...lib.classifyFileName(hwpx), ...lib.inspectFormFile(hwpx) });
  const calls = [];
  const printedText = [];
  const hancom = async (action, args) => {
    calls.push(action);
    if (action === 'convert') {
      fs.writeFileSync(args.output, hwpxWith([P(T('이름: '), F('이름'))]));
      return { ok: true };
    }
    printedText.push({ file: args.file, text: textOf(fs.readFileSync(args.file)), copies: args.copies });
    return printedText.length === 1 ? { ok: true } : { ok: false, error: '[가짜] 프린터 오류' };
  };
  const results = await job.printForms(L, [a, b], INPUT, { copies: 2, hancom });
  assert.deepStrictEqual(calls, ['convert', 'print', 'print']);
  assert.deepStrictEqual(results.map((r) => r.ok), [true, false]);
  assert.strictEqual(results[1].error, '[가짜] 프린터 오류');
  assert.deepStrictEqual(printedText.map((p) => p.text), ['이름: 홍길동\n', '성명 홍길동\n']);
  assert.ok(printedText.every((p) => !fs.existsSync(p.file) && p.copies === 2), '임시 파일이 남았습니다');
  assert.ok(L.get(a.id).hwpx, '변환본을 보관함에 넣음');
  const again = await job.printForms(L, [L.get(a.id)], INPUT, { hancom });
  assert.ok(again[0].ok === false && calls.filter((c) => c === 'convert').length === 1, '두 번째부터는 변환하지 않는다');
  const noHancom = await job.printForms(L, [lib.openLibrary(path.join(ldir, 'forms')).get(b.id)], INPUT, { hancom: async () => ({ ok: false, error: '한글 자동 조작은 한글이 설치된 Windows에서만 할 수 있습니다.' }) });
  assert.ok(noHancom[0].error.includes('Windows'));
});

/* ---------- 서류 출력 규칙(print-plan.js) ---------- */

const entry = (id, over = {}) => ({ id, title: id, agency: '양주시청', kind: '긴급복지(국가형)', applicant: '세대주', favorite: false, documents: [{ mark: '[서식 1호]', title: '현장확인서' }], fields: [], mapping: {}, ...over });

test('출력 규칙: 기관·구분에 맞는 서식(공통 포함), 지원 종류 → ★ → 이름 순, 지원 종류별 묶음', () => {
  const list = [
    entry('나', { kind: '경기도형 긴급복지' }),
    entry('가'),
    entry('다', { favorite: true }),
    entry('세대원용', { applicant: '세대원' }),
    entry('공통', { applicant: '공통', kind: '암환자 의료비' }),
    entry('다른 기관', { agency: '노원구청' }),
  ];
  const shown = plan.entriesFor(list, '양주시청', '세대주');
  assert.deepStrictEqual(shown.map((e) => e.id), ['다', '가', '나', '공통']);
  assert.deepStrictEqual(plan.groupByKind(shown).map((g) => [g.kind, g.items.length]), [['긴급복지(국가형)', 2], ['경기도형 긴급복지', 1], ['암환자 의료비', 1]]);
  assert.deepStrictEqual(plan.entriesFor(list, '양주시청', '세대원').map((e) => e.id), ['세대원용', '공통']);
});

test('출력 규칙: 필요한 입력·비어 있는 칸·출력 때 입력·설정의 담당자', () => {
  const e = entry('암', { fields: ['이름', '주민등록번호', '주질병', '담당자', '진료기간', '보호자', '기관'].map((name) => ({ name, count: 1 })), mapping: { 보호자: { type: 'ask' }, 기관: { type: 'fixed', text: '[예시]' } } });
  assert.deepStrictEqual(plan.neededInputs([e]), ['name', 'rrn', 'careStart', 'careEnd', 'disease', 'staffName']);
  assert.deepStrictEqual(plan.askFieldNames([e, e]), ['보호자']);
  assert.deepStrictEqual(plan.missingFor(e, { name: '홍길동' }, {}), ['주민등록번호', '진료 시작일', '주질병(상병명)', '담당자', '보호자']);
  assert.deepStrictEqual(plan.missingFor(e, { ...INPUT, careStart: '2026-01-02' }, { 보호자: '[x]' }), []);
  assert.deepStrictEqual(plan.missingStaff([e], { staffName: '' }), ['담당자']);
  assert.deepStrictEqual(plan.missingStaff([e], { staffName: '[담당자]' }), []);
});

test('출력 규칙: 출력 전 확인(필수 칸·주민등록번호 모양은 막고, 빈칸은 알림), 요약, 저장 파일 이름', () => {
  const e = entry('서식', { fields: [{ name: '이름', count: 1 }, { name: '4', count: 1 }, { name: '연락처', count: 1 }] });
  assert.deepStrictEqual(plan.checkBeforeOutput([e], {}, {}), { block: '이름, 주민등록번호을(를) 넣어 주세요.', focus: 'name', missing: [] });
  assert.deepStrictEqual(plan.checkBeforeOutput([e], { name: 'x', rrn: '123' }, {}), { block: '주민등록번호를 확인하세요: 주민등록번호는 숫자 13자리입니다', focus: 'rrn', missing: [] });
  assert.deepStrictEqual(plan.checkBeforeOutput([e], { name: 'x', rrn: '900101-1234567' }, {}), { block: '', focus: null, missing: ['연락처'] });
  assert.deepStrictEqual(plan.summaryOf([], {}, {}, 1), { title: '출력할 서류를 고르세요', detail: '' });
  assert.deepStrictEqual(plan.summaryOf([e, e], { name: 'x', rrn: '900101-1234567', phone: '1' }, {}, 2), { title: '서식 2개 · 서류 약 2종 × 2부', detail: '모든 칸이 채워집니다' });
  assert.strictEqual(plan.filledFileName(entry('a/b:c')), '양주시청_a_b_c(세대주).hwpx');
  assert.strictEqual(plan.filledFileName(entry('암', { applicant: '공통', agency: '공통 서식' })), '공통 서식_암.hwpx');
});

test('설정: 없으면 기본값, 부수는 1~99, 망가진 파일은 기본값과 오류', () => {
  const sdir = fs.mkdtempSync(path.join(dir, 'set-'));
  assert.deepStrictEqual(loadSettings(sdir).settings, { staffName: '', staffTitle: '', staffPhone: '', copies: 1, printer: '' });
  saveSettings(sdir, { staffName: ' [담당자] ', copies: '250', extra: 'x' });
  assert.deepStrictEqual(loadSettings(sdir).settings, { staffName: '[담당자]', staffTitle: '', staffPhone: '', copies: 99, printer: '' });
  fs.writeFileSync(path.join(sdir, 'settings.json'), '{');
  assert.ok(loadSettings(sdir).error && loadSettings(sdir).settings.copies === 1);
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
  fs.rmSync(dir, { recursive: true, force: true });
  console.log(`\n[autofill] ${passed}/${passed + failed} 통과`);
  process.exit(failed ? 1 : 0);
})();
