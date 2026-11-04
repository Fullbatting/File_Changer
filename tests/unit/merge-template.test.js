// tests/unit/merge-template.test.js
// 메일머지 템플릿(merge-template.js): Word·한글 {{변수}} 찾기·바꾸기, 결과 파일 이름, 빈 값 검사. node tests/unit/merge-template.test.js

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const PizZip = require('pizzip');
const merge = require('../../merge-template.js');

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

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'merge-template-'));
const NS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"';
const P = (runs) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0">${runs}</hp:p>`;
const R = (t, id = 0) => `<hp:run charPrIDRef="${id}"><hp:t>${t}</hp:t></hp:run>`;

function writeHwpx(name, sections) {
  const zip = new PizZip();
  zip.file('mimetype', 'application/hwp+zip');
  sections.forEach((body, i) => zip.file(`Contents/section${i}.xml`, `<?xml version="1.0" encoding="UTF-8"?><hs:sec ${NS}>${body}</hs:sec>`));
  const file = path.join(dir, name);
  fs.writeFileSync(file, zip.generate({ type: 'nodebuffer' }));
  return file;
}

function writeDocx(name, bodyXml, { template = false, header = '' } = {}) {
  const zip = new PizZip();
  const main = template ? 'template' : 'document';
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${main}.main+xml"/>${header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}</Types>`);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document>`);
  if (header) zip.file('word/header1.xml', `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${header}</w:hdr>`);
  const file = path.join(dir, name);
  fs.writeFileSync(file, zip.generate({ type: 'nodebuffer' }));
  return file;
}

const sectionsOf = (buffer) => {
  const zip = new PizZip(buffer);
  return Object.keys(zip.files)
    .filter((n) => /section\d+\.xml$/.test(n))
    .sort()
    .map((n) => zip.file(n).asText());
};

const HWPX = writeHwpx('tpl.hwpx', [
  [
    P(R('수신: {{이름}} 님 ({{ 이름 }})')),
    P(R('주소: {{주', 1) + R('소', 2) + R('}} 끝', 3)),
    P(R('A &amp; B {{회사명}}')),
    P(R('{{#목록}}반복{{/목록}} {{.}}')),
    P(R('비고: [{{비고}}]')),
  ].join(''),
  P(R('둘째 구역 {{이름}} {{메모}}')),
]);

test('형식: .docx/.dotx → docx, .hwpx/.hwpt → hwpx, .hwp·그 밖은 안내와 함께 거부', () => {
  assert.strictEqual(merge.templateFormat('a.DOTX'), 'docx');
  assert.strictEqual(merge.templateFormat('a.hwpt'), 'hwpx');
  assert.throws(() => merge.templateFormat('a.hwp'), /다른 이름으로 저장 > HWPX/);
  assert.throws(() => merge.templateFormat('a.pdf'), /지원하지 않는 템플릿 형식/);
  assert.ok(merge.isControlToken('#목록') && merge.isControlToken('/목록') && merge.isControlToken('.') && !merge.isControlToken('이름'));
});

test('한글: 변수는 처음 나온 순서대로 한 번씩(구역을 넘어서도), 여러 런으로 나뉜 변수도, 제어 태그는 뺀다', () => {
  assert.deepStrictEqual(merge.openTemplate(HWPX).variables, ['이름', '주소', '회사명', '비고', '메모']);
});

test('한글: 값으로 바꾸기 — 나뉜 런은 서식 유지, 특수문자·줄바꿈, 없는 변수·제어 태그는 지움, 값의 앞뒤 공백 유지', () => {
  const [s0, s1] = sectionsOf(merge.openTemplate(HWPX).render({ 이름: '홍길동', 주소: '○○시 & <구>', 회사명: ' [예시] ', 비고: '첫째\n둘째', 메모: '' }));
  assert.ok(s0.includes('<hp:t>수신: 홍길동 님 (홍길동)</hp:t>'), s0);
  assert.ok(s0.includes('<hp:t>주소: ○○시 &amp; &lt;구&gt;</hp:t></hp:run><hp:run charPrIDRef="2"><hp:t></hp:t></hp:run><hp:run charPrIDRef="3"><hp:t> 끝</hp:t>'), s0);
  assert.ok(s0.includes('<hp:t>A &amp; B  [예시] </hp:t>'), s0);
  assert.ok(s0.includes('<hp:t>반복 </hp:t>'), s0);
  assert.ok(s0.includes('<hp:t>비고: [첫째<hp:lineBreak/>둘째]</hp:t>'), s0);
  assert.ok(s1.includes('<hp:t>둘째 구역 홍길동 </hp:t>'), s1);
});

test('한글: 같은 템플릿으로 여러 번 만들어도 서로 섞이지 않는다(일괄 생성)', () => {
  const t = merge.openTemplate(HWPX);
  const a = sectionsOf(t.render({ 이름: '갑' }))[1];
  const b = sectionsOf(t.render({ 이름: '을' }))[1];
  assert.ok(a.includes('구역 갑 ') && b.includes('구역 을 ') && !b.includes('갑'));
});

test('템플릿 열기: 같은 파일은 다시 읽지 않고, 파일이 바뀌면 새로 읽는다', () => {
  const file = writeHwpx('cache.hwpx', [P(R('{{가}}'))]);
  const first = merge.openTemplate(file);
  assert.strictEqual(merge.openTemplate(file), first);
  writeHwpx('cache.hwpx', [P(R('{{나}} {{다}}'))]);
  const later = new Date(Date.now() + 5000);
  fs.utimesSync(file, later, later);
  assert.deepStrictEqual(merge.openTemplate(file).variables, ['나', '다']);
});

test('Word: 본문 → 머리글 순서로 변수, 나뉜 <w:t>도, 값이 없으면 빈칸, .dotx 결과는 문서 형식으로', () => {
  const body = '<w:p><w:r><w:t>{{이</w:t></w:r><w:r><w:t>름}} {{부서}}</w:t></w:r></w:p><w:p><w:r><w:t>{{#목록}}x{{/목록}}</w:t></w:r></w:p>';
  const file = writeDocx('tpl.dotx', body, { template: true, header: '<w:p><w:r><w:t>{{회사명}} {{이름}}</w:t></w:r></w:p>' });
  const t = merge.openTemplate(file);
  assert.deepStrictEqual(t.variables, ['이름', '부서', '회사명']);
  const zip = new PizZip(t.render({ 이름: '홍길동 & <1>', 회사명: '[예시]' }));
  const doc = zip.file('word/document.xml').asText();
  assert.ok(doc.includes('홍길동 &amp; &lt;1&gt;') && !doc.includes('undefined') && !doc.includes('{{'), doc);
  assert.ok(zip.file('word/header1.xml').asText().includes('[예시] 홍길동'));
  assert.ok(zip.file('[Content_Types].xml').asText().includes('wordprocessingml.document.main+xml'));
  assert.throws(() => merge.docxVariables(new PizZip().file('a.txt', 'x').generate({ type: 'nodebuffer' })), /word\/document\.xml 없음/);
});

test('파일 이름: 못 쓰는 글자·끝의 점·예약 이름, 빈 값은 "행N", 대소문자만 다른 이름도 겹침으로 (2)', () => {
  assert.strictEqual(merge.sanitizeFilename('a/b:c*?.'), 'a_b_c__');
  assert.strictEqual(merge.sanitizeFilename('con'), '_con');
  const rows = [{ 이름: '홍길동' }, { 이름: '' }, { 이름: '홍길동' }, { 이름: 'Kim' }, { 이름: 'kim' }];
  assert.deepStrictEqual(merge.outputFilenames(rows, '이름', 'hwpx'), ['홍길동.hwpx', '행2.hwpx', '홍길동(2).hwpx', 'Kim.hwpx', 'kim(2).hwpx']);
  assert.deepStrictEqual(merge.outputFilenames(rows.slice(0, 2), '', 'docx'), ['행1.docx', '행2.docx']);
  assert.strictEqual(merge.unnamedRowCount(rows, '이름'), 1);
  assert.strictEqual(merge.unnamedRowCount(rows, ''), 5);
  assert.strictEqual(merge.defaultFilenameColumn(['번호', '성명', '이름']), '성명');
});

test('일괄 생성: 행 값은 화면에 보이는 글자로, 매핑 안 된 변수·빈 값 개수', () => {
  const rows = [{ A: '홍길동', B: 1500 }, { A: '', B: null }];
  const mapping = { 이름: 'A', 금액: 'B', 비고: '' };
  assert.deepStrictEqual(merge.rowData(rows[0], ['이름', '금액', '비고'], mapping), { 이름: '홍길동', 금액: '1500', 비고: '' });
  assert.deepStrictEqual(merge.missingValueCounts(rows, ['이름', '금액', '비고'], mapping), { 이름: 1, 금액: 1, 비고: 2 });
});

fs.rmSync(dir, { recursive: true, force: true });
console.log(`\n[merge-template] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
