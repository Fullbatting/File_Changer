// tests/unit/table-reader.test.js
// 파일 인식(헤더 행·빈 머리글 열·1904 날짜·인코딩·내용 형식) 단위 테스트: node tests/unit/table-reader.test.js
// 날짜 검사는 현지 시각 기준이라 시간대와 무관하다.

const assert = require('assert');
const XLSX = require('xlsx');
const iconv = require('iconv-lite');
const t = require('../../table-reader.js');

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

const sheet = (rows) => XLSX.utils.aoa_to_sheet(rows.map((r) => r.map((v) => (v === '' ? null : v))));
const ymd = (d) => `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()} ${d.getHours()}:${d.getMinutes()}`;

test('헤더 판정: 머리글 칸 하나가 비어 데이터 행이 더 많이 채워져도 헤더 행을 고른다', () => {
  const r = t.buildTableFromSheet(sheet([['이름', '전화', ''], ['홍길동', '010', '서울'], ['김영희', '011', '부산']]), 'S');
  assert.strictEqual(r.headerRowIndex, 0);
  assert.deepStrictEqual(r.rows.map((x) => x['이름']), ['홍길동', '김영희']);
});

test('빈 머리글 열: 값이 있으면 버리지 않고 "(이름 없음) C열"로 넣고 알린다', () => {
  const r = t.buildTableFromSheet(sheet([['이름', '전화', ''], ['홍길동', '010', '서울']]), 'S');
  assert.deepStrictEqual(r.headers, ['이름', '전화', '(이름 없음) C열']);
  assert.strictEqual(r.rows[0]['(이름 없음) C열'], '서울');
  assert.deepStrictEqual(r.report.unnamedColumns, ['(이름 없음) C열']);
  assert.ok(t.describeRecognition(r.report).some((i) => i.warn && i.text.includes('(이름 없음) C열')));
});

test('빈 머리글 열: 값도 없으면 뺀다', () => {
  const r = t.buildTableFromSheet(sheet([['이름', '', '부서'], ['홍길동', '', '영업']]), 'S');
  assert.deepStrictEqual(r.headers, ['이름', '부서']);
  assert.deepStrictEqual(r.report.droppedEmptyColumns, ['B']);
});

test('헤더 판정: 제목 한 줄·작성 정보 줄 아래의 헤더를 고르고 건너뛴 행을 보고', () => {
  const r = t.buildTableFromSheet(
    sheet([
      ['2026 직원 명부', '', '', ''],
      ['작성일', '2026-01-05', '', ''],
      ['', '', '', ''],
      ['이름', '사번', '부서', '입사일'],
      ['홍길동', '20230001', '영업팀', '2026-01-15'],
    ]),
    'S'
  );
  assert.strictEqual(r.headerRowIndex, 3);
  assert.deepStrictEqual(r.report.skippedRows.map((x) => x.rowNumber), [1, 2]);
  assert.strictEqual(r.rows.length, 1);
  assert.deepStrictEqual(r.rowNumbers, [5]);
});

test('헤더 판정: 두 줄 머리글이면 아래 줄(세로 병합으로 같은 이름이 있거나 바로 아래부터 데이터)', () => {
  // 한글 문서 표: 이름·부서는 두 줄로 세로 병합, 연락처는 전화번호·이메일 두 칸 묶음
  const hwp = t.buildTableFromSheet(sheet([['2026년 하반기 직원 명단', '', '', ''], ['이름', '연락처', '', '부서'], ['이름', '전화번호', '이메일', '부서'], ['홍길동', '010-1111-2222', 'hong@example.com', '영업팀']]), 'S');
  assert.strictEqual(hwp.headerRowIndex, 2);
  assert.deepStrictEqual(hwp.headers, ['이름', '전화번호', '이메일', '부서']);
  // 병합 없이 묶음 이름만 있는 위 줄, 아래는 숫자가 섞인 데이터
  const group = t.buildTableFromSheet(sheet([['인적사항', '', '급여', ''], ['이름', '부서', '기본급', '수당'], ['홍길동', '영업', 100, 20]]), 'S');
  assert.strictEqual(group.headerRowIndex, 1);
  // 글자만 있는 데이터가 이어지면 위 줄(진짜 헤더)에 머문다
  const textData = t.buildTableFromSheet(sheet([['이름', '부서', ''], ['홍길동', '영업팀', '서울'], ['김영희', '개발팀', '부산']]), 'S');
  assert.strictEqual(textData.headerRowIndex, 0);
});

test('헤더 판정: 숫자 컬럼 이름(연도)만 있는 표는 예전 방식(가장 많이 채워진 행)으로', () => {
  const r = t.buildTableFromSheet(sheet([['보고서'], [2024, 2025, 2026], [10, 20, 30]]), 'S');
  assert.strictEqual(r.headerRowIndex, 1);
});

test('헤더 행 직접 고르기: 자동과 다른 행을 쓰고 report에 표시', () => {
  const ws = sheet([['구분', '', ''], ['이름', '전화', '주소'], ['홍길동', '010', '서울']]);
  const auto = t.buildTableFromSheet(ws, 'S');
  assert.strictEqual(auto.headerRowIndex, 1);
  const manual = t.buildTableFromSheet(ws, 'S', { headerRowIndex: 0 });
  assert.strictEqual(manual.headerRowIndex, 0);
  assert.strictEqual(manual.report.headerAuto, false);
  assert.strictEqual(manual.report.autoHeaderRowNumber, 2);
  assert.strictEqual(manual.rows.length, 2);
  const bad = t.buildTableFromSheet(ws, 'S', { headerRowIndex: 99 });
  assert.strictEqual(bad.headerRowIndex, 1);
});

test('헤더 후보 목록: 맨 위 10행, 미리보기 글자와 헤더로 보이는지', () => {
  const r = t.buildTableFromSheet(sheet([['명단'], ['이름', '부서'], ['홍길동', '영업']]), 'S');
  const c = r.report.headerCandidates;
  assert.deepStrictEqual(c.map((x) => x.rowNumber), [1, 2, 3]);
  assert.strictEqual(c[1].likely, true);
  assert.strictEqual(c[1].preview, '이름, 부서');
});

test('가로 병합 머리글: 뒤 칸이 앞 칸 이름을 이어 받는다', () => {
  const ws = sheet([['이름', '주소', ''], ['홍길동', '서울', '종로구']]);
  ws['!merges'] = [{ s: { r: 0, c: 1 }, e: { r: 0, c: 2 } }];
  const r = t.buildTableFromSheet(ws, 'S');
  assert.deepStrictEqual(r.headers, ['이름', '주소', '주소_1']);
  assert.strictEqual(r.rows[0]['주소_1'], '종로구');
});

function dateWorkbook(date1904) {
  const ws = XLSX.utils.aoa_to_sheet([['날짜', '일시'], [0, 0]]);
  // 2026-01-15 와 2026-01-15 13:45 — 1904 체계는 같은 날짜의 일련번호가 1462 작다
  const base = 46037 - (date1904 ? 1462 : 0);
  ws.A2 = { t: 'n', v: base, z: 'yyyy-mm-dd' };
  ws.B2 = { t: 'n', v: base + (13 * 60 + 45) / 1440, z: 'yyyy-mm-dd hh:mm' };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'S');
  if (date1904) wb.Workbook = { ...(wb.Workbook || {}), WBProps: { date1904: true } };
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return XLSX.read(buffer, { type: 'buffer', cellNF: true });
}

test('dense 모드(행 배열)로 읽은 시트도 같은 표: 제목 행·B2 시작·병합·형식', () => {
  const ws = sheet([['', '', ''], ['', '명단', ''], ['', '이름', '급여'], ['', '홍길동', 1500000], ['', '', 2000]]);
  ws.C4.z = '#,##0';
  ws['!merges'] = [{ s: { r: 3, c: 1 }, e: { r: 4, c: 1 } }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'S');
  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const sparse = t.buildTableFromSheet(XLSX.read(buffer, { type: 'buffer', cellNF: true }).Sheets.S, 'S');
  const denseSheet = XLSX.read(buffer, { type: 'buffer', cellNF: true, dense: true }).Sheets.S;
  assert.ok(denseSheet['!data'], 'dense 시트가 아님');
  const dense = t.buildTableFromSheet(denseSheet, 'S');
  assert.deepStrictEqual(dense.rows, sparse.rows);
  assert.deepStrictEqual(dense.rows, [{ 이름: '홍길동', 급여: 1500000 }, { 이름: '홍길동', 급여: 2000 }]);
  assert.deepStrictEqual(dense.rowNumbers, [4, 5]);
  assert.deepStrictEqual(dense.numberFormats, { 급여: '#,##0' });
  assert.strictEqual(t.cellAt(denseSheet, 2, 1).v, '이름');
});

test('1900 날짜 체계: 날짜와 날짜+시각', () => {
  const wb = dateWorkbook(false);
  assert.strictEqual(t.isDate1904(wb), false);
  const r = t.buildTableFromSheet(wb.Sheets.S, 'S', { date1904: t.isDate1904(wb) });
  assert.strictEqual(ymd(r.rows[0]['날짜']), '2026-1-15 0:0');
  assert.strictEqual(ymd(r.rows[0]['일시']), '2026-1-15 13:45');
});

test('1904 날짜 체계: 4년 어긋나지 않고 같은 날짜로', () => {
  const wb = dateWorkbook(true);
  assert.strictEqual(t.isDate1904(wb), true);
  const r = t.buildTableFromSheet(wb.Sheets.S, 'S', { date1904: true });
  assert.strictEqual(ymd(r.rows[0]['날짜']), '2026-1-15 0:0');
  assert.strictEqual(ymd(r.rows[0]['일시']), '2026-1-15 13:45');
  assert.ok(t.describeRecognition(r.report).some((i) => i.text.includes('1904')));
});

test('내용 형식 확인: zip·ole·html·글자, 확장자와 다르면 안내', () => {
  assert.strictEqual(t.sniffContentKind(Buffer.from([0x50, 0x4b, 0x03, 0x04, 1, 2])), 'zip');
  assert.strictEqual(t.sniffContentKind(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 1, 2])), 'ole');
  assert.strictEqual(t.sniffContentKind(Buffer.from('<html><table><tr><td>1</td></tr></table></html>')), 'html');
  assert.strictEqual(t.sniffContentKind(Buffer.from('이름,부서\n홍길동,영업')), 'text');
  assert.strictEqual(t.describeContentMismatch('.xlsx', 'zip'), '');
  assert.ok(t.describeContentMismatch('.xls', 'html').includes('HTML 표'));
  assert.ok(t.describeContentMismatch('.xlsx', 'text').includes('CSV'));
});

test('글자 인코딩: BOM·UTF-8·CP949 자동 판별과 직접 지정, 깨진 글자 수', () => {
  const text = '이름,부서\n홍길동,영업팀';
  assert.strictEqual(t.decodeTextBuffer(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)])).encoding, 'utf-8-bom');
  assert.strictEqual(t.decodeTextBuffer(Buffer.from(text)).encoding, 'utf-8');
  const cp = iconv.encode(text, 'cp949');
  const auto = t.decodeTextBuffer(cp);
  assert.strictEqual(auto.encoding, 'cp949');
  assert.strictEqual(auto.text, text);
  const forced = t.decodeTextBuffer(cp, 'utf-8');
  assert.strictEqual(forced.auto, false);
  assert.ok(forced.brokenChars > 0);
});

console.log(`\n[table-reader] ${passed}/${passed + failed} 통과`);
process.exit(failed ? 1 : 0);
