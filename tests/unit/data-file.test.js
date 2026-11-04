// tests/unit/data-file.test.js
// 데이터 파일 읽기(data-file.js)·동의어 사전(synonyms.js)·파일 도우미(file-utils.js). node tests/unit/data-file.test.js

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const PizZip = require('pizzip');
const dataFile = require('../../data-file.js');
const synonyms = require('../../synonyms.js');
const { createFileCache, friendlyErrorMessage, safeReadFileSync } = require('../../file-utils.js');

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

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'data-file-'));
const NS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"';
const P = (inner) => `<hp:p id="0">${inner}</hp:p>`;
/** 칸: lines는 문단 글자들(글자 안 "{nb}"는 묶음 빈칸, "{tab}"은 탭) */
const tc = (row, col, lines, { rowSpan = 1, colSpan = 1, inner = '' } = {}) =>
  `<hp:tc><hp:subList>${[]
    .concat(lines)
    .map((t, i) => P(`<hp:run charPrIDRef="0"><hp:t>${t.replace(/\{nb\}/g, '<hp:nbSpace/>').replace(/\{tab\}/g, '<hp:tab/>')}</hp:t>${i === 0 ? inner : ''}</hp:run>`))
    .join('')}</hp:subList><hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="${colSpan}" rowSpan="${rowSpan}"/></hp:tc>`;
const tbl = (trs) => `<hp:tbl rowCnt="${trs.length}" colCnt="9">${trs.map((tr) => `<hp:tr>${tr.join('')}</hp:tr>`).join('')}</hp:tbl>`;

const nested = tbl([[tc(0, 0, '안쪽')]]);
const staffXml = [
  P(`<hp:run>${tbl([[tc(0, 0, '담당'), tc(0, 1, '팀장')], [tc(1, 0, ''), tc(1, 1, '')]])}<hp:t/></hp:run>`),
  P(`<hp:run>${tbl([[tc(0, 0, '')]])}<hp:t/></hp:run>`),
  P(
    `<hp:run>${tbl([
      [tc(0, 0, '이름', { rowSpan: 2 }), tc(0, 1, '연락처', { colSpan: 2 }), tc(0, 3, '부서', { rowSpan: 2, inner: nested })],
      [tc(1, 1, '전화'), tc(1, 2, '이메일')],
      [tc(2, 0, '홍{nb}길동'), tc(2, 1, '010-0000-0000'), tc(2, 2, 'a&amp;b@example.com'), tc(2, 3, ['개발', '2팀'])],
      [tc(3, 0, ' 김{tab}영희 '), tc(3, 1, ''), tc(3, 2, ''), tc(3, 3, '영업')],
    ])}<hp:t/></hp:run>`
  ),
].join('');
const zip = new PizZip();
zip.file('mimetype', 'application/hwp+zip');
zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?><hs:sec ${NS}>${staffXml}</hs:sec>`);
const STAFF = path.join(dir, 'staff.hwpx');
fs.writeFileSync(STAFF, zip.generate({ type: 'nodebuffer' }));

test('한글 표: 문서 순서(바깥 표 다음에 그 안의 표)·중첩 표 글자는 따로, 병합(세로는 채움·가로는 비움), 특수 빈칸·탭은 빈칸, 문단은 줄바꿈, &amp; 풀기', () => {
  const tables = dataFile.readHwpxTables(fs.readFileSync(STAFF));
  assert.strictEqual(tables.length, 4);
  assert.deepStrictEqual(tables[3].rows, [['안쪽']], '칸 안에 든 표는 바깥 표 다음에 따로');
  assert.deepStrictEqual(tables[2].rows, [
    ['이름', '연락처', '', '부서'],
    ['이름', '전화', '이메일', '부서'],
    ['홍 길동', '010-0000-0000', 'a&b@example.com', '개발\n2팀'],
    ['김 영희', '', '', '영업'],
  ]);
});

test('한글 표 → 워크북: 빈 표는 빼고 "표 N", 행이 가장 많은 표를 먼저, 찾은 표 요약 알림(새로 읽을 때만)', () => {
  const notes = [];
  const wb = dataFile.readWorkbook(STAFF, { onNote: (n) => notes.push(n) });
  dataFile.readWorkbook(STAFF, { onNote: (n) => notes.push(n) });
  assert.deepStrictEqual(wb.SheetNames, ['표 1', '표 2', '표 3']);
  assert.deepStrictEqual(dataFile.sheetChoices(wb), { names: ['표 1', '표 2', '표 3'], unit: '표', selected: '표 2' });
  assert.strictEqual(notes.length, 1);
  assert.ok(notes[0].startsWith('한글 문서에서 표 3개를 찾았습니다: 표 1(2행×2열: 담당, 팀장)'), notes[0]);
  assert.throws(() => dataFile.workbookFromTables([{ rows: [['']] }]), /표를 찾을 수 없습니다/);
});

test('한글 표 → 헤더·행: 2단 머리글을 합쳐 열 이름을 만든다', () => {
  const table = dataFile.readSheetTable(STAFF, '표 2', { displayText: true });
  assert.strictEqual(table.rows.length, 2);
  assert.ok(table.textSource);
  assert.strictEqual(table.rows[0][table.headers[0]], '홍 길동');
});

test('CSV: 내용으로 판단하고 글자 그대로(앞자리 0 유지), 파일이 바뀌면 다시 읽는다', () => {
  const csv = path.join(dir, 'people.csv');
  fs.writeFileSync(csv, '이름,우편번호\n홍길동,03187\n');
  assert.strictEqual(dataFile.readSheetTable(csv, null).rows[0]['우편번호'], '03187');
  fs.writeFileSync(csv, '이름,우편번호\n김영희,01234\n');
  const later = new Date(Date.now() + 5000);
  fs.utimesSync(csv, later, later);
  assert.strictEqual(dataFile.readSheetTable(csv, null).rows[0]['이름'], '김영희');
  assert.throws(() => dataFile.readSheetTable(csv, '없는 시트'), /시트를 찾을 수 없습니다 \(없는 시트\)/);
});

test('손상된 한글 문서: 해석할 수 없다는 안내', () => {
  const broken = new PizZip();
  broken.file('Contents/section0.xml', '<hs:sec><hp:p>');
  assert.throws(() => dataFile.readHwpxTables(broken.generate({ type: 'nodebuffer' })), /한글 문서 본문을 해석할 수 없습니다/);
  assert.throws(() => dataFile.readHwpxTables(new PizZip().file('a', 'b').generate({ type: 'nodebuffer' })), /Contents\/section\*\.xml 없음/);
});

test('동의어 사전: 없으면 기본 사전을 저장, 망가진 파일은 기본 사전 + 오류(파일은 그대로), 글 ↔ 그룹', () => {
  const sdir = fs.mkdtempSync(path.join(dir, 'syn-'));
  const first = synonyms.loadSynonymGroups(sdir);
  assert.deepStrictEqual(first.groups, synonyms.DEFAULT_SYNONYM_GROUPS);
  assert.ok(fs.existsSync(path.join(sdir, 'synonym-groups.json')));
  synonyms.saveSynonymGroups(sdir, [['월급', '급여'], ['혼자']]);
  assert.deepStrictEqual(synonyms.loadSynonymGroups(sdir).groups, [['월급', '급여']], '이름이 하나뿐인 그룹은 뺀다');
  fs.writeFileSync(path.join(sdir, 'synonym-groups.json'), '[[broken');
  const broken = synonyms.loadSynonymGroups(sdir);
  assert.ok(broken.error && broken.groups.length === synonyms.DEFAULT_SYNONYM_GROUPS.length);
  assert.strictEqual(fs.readFileSync(path.join(sdir, 'synonym-groups.json'), 'utf8'), '[[broken');
  assert.deepStrictEqual(synonyms.textToGroups('이름, 성명 ,\n혼자\n\n a ,b'), [['이름', '성명'], ['a', 'b']]);
  assert.strictEqual(synonyms.groupsToText([['이름', '성명'], ['a', 'b']]), '이름, 성명\na, b');
});

test('파일 도우미: 캐시는 파일이 바뀌면 새로, 개수 제한은 먼저 넣은 것부터 버림 · 오류 문구에 경로', () => {
  const f = path.join(dir, 'c.txt');
  fs.writeFileSync(f, '1');
  const cache = createFileCache(2);
  let loads = 0;
  const load = () => (loads += 1);
  cache.get(f, load);
  cache.get(f, load);
  assert.strictEqual(loads, 1);
  const later = new Date(Date.now() + 9000);
  fs.utimesSync(f, later, later);
  cache.get(f, load);
  assert.strictEqual(loads, 2);
  cache.get(f, load, 'k2');
  cache.get(f, load, 'k3');
  assert.strictEqual(cache.size, 2);
  let err;
  try {
    safeReadFileSync(path.join(dir, '없음.txt'));
  } catch (e) {
    err = e;
  }
  assert.ok(friendlyErrorMessage(err).startsWith('파일을 찾을 수 없습니다 (') && err.code === 'ENOENT');
  assert.ok(friendlyErrorMessage({ code: 'EBUSY', path: 'x.xlsx' }).includes('다른 프로그램'));
});

fs.rmSync(dir, { recursive: true, force: true });
console.log(`\n[data-file] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
