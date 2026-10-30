// Fixtures for the full feature audit. Excel files are written as Excel stores cells (typed values,
// number formats, merges, used ranges not starting at A1), so the app's read path is exercised realistically.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const iconv = require('iconv-lite');
const D = process.env.LEGACY_FIX;

const s = (v) => ({ t: 's', v });
const n = (v, z) => (z ? { t: 'n', v, z } : { t: 'n', v });
function writeXlsx(file, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([name, ws]) => XLSX.utils.book_append_sheet(wb, ws, name));
  fs.writeFileSync(path.join(D, file), XLSX.write(wb, { type: 'buffer', bookType: file.endsWith('.xls') ? 'biff8' : 'xlsx' }));
}

// Source A: table starts at B2, header cells with stray spaces, duplicate "비고", typed values:
// date (serial 46037 = 2026-01-15), number 1500000, percent 0.15, zero-padded code 3187 shown "03187",
// text digits "20230001" (사번 stored as text), a vertically merged 부서 cell.
const src = {
  '!ref': 'B2:J5',
  B2: s('이름 '), C2: s(' 사번'), D2: s('입사일'), E2: s('급여'), F2: s('비율'), G2: s('우편번호'), H2: s('비고'), I2: s('비고'), J2: s('부서'),
  B3: s('홍길동'), C3: s('20230001'), D3: n(46037, 'yyyy-mm-dd'), E3: n(1500000, '#,##0'), F3: n(0.15, '0%'), G3: n(3187, '00000'), H3: s('첫째'), I3: s('둘째'), J3: s('영업팀'),
  B4: s('김영희'), C4: s('20230002'), D4: n(46072.5, 'yyyy-mm-dd hh:mm'), E4: n(2250000, '#,##0'), F4: n(0.2, '0%'), G4: n(4321, '00000'), H4: s(''), I4: s('셋째'),
  B5: s('홍길동'), C5: s('20230001'), D5: n(46037, 'yyyy-mm-dd'), E5: n(1500000, '#,##0'), F5: n(0.15, '0%'), G5: n(3187, '00000'), H5: s('첫째'), I5: s('둘째'), J5: s('영업팀'),
  '!merges': [{ s: { r: 2, c: 9 }, e: { r: 3, c: 9 } }],
};
// Second sheet to exercise the sheet selector.
const other = { '!ref': 'A1:B2', A1: s('항목'), B1: s('값'), A2: s('x'), B2: s('y') };
writeXlsx('source.xlsx', [['직원', src], ['기타', other]]);

// Target B: title row, then headers incl. synonym (성명), jamo-typo (승명 -> 성명? no: 입사일 vs 입사릴), duplicate 비고.
const tgt = {
  '!ref': 'A1:H2',
  A1: s('2026년 직원 명부'),
  A2: s('성명'), B2: s('사번'), C2: s('입사릴'), D2: s('급여'), E2: s('비율'), F2: s('우편번호'), G2: s('비고'), H2: s('비고'),
};
writeXlsx('target.xlsx', [['명부', tgt]]);
writeXlsx('target.xls', [['명부', tgt]]);

// CSV in CP949 (Windows Korean Excel default) with leading zeros and a date.
fs.writeFileSync(path.join(D, 'source_cp949.csv'), iconv.encode('이름,전화번호,우편번호,가입일\r\n홍길동,01012345678,03187,2026-01-15\r\n김영희,01098765432,04321,2026-02-20\r\n', 'cp949'));
// UTF-8 BOM CSV
fs.writeFileSync(path.join(D, 'source_bom.csv'), Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('이름,전화번호\n이철수,01055556666\n')]));
const csvTarget = { '!ref': 'A1:D1', A1: s('성명'), B1: s('연락처'), C1: s('우편번호'), D1: s('가입일') };
writeXlsx('target_csv.xlsx', [['명부', csvTarget]]);

// Mail-merge data: 250 rows (exercises the yielding loop + progress logs), names with duplicates in different
// case, a Windows-reserved name, a trailing dot, a date column, a multi-line address, a percent.
const mm = { '!ref': 'A1:E251', A1: s('이름'), B1: s('가입일'), C1: s('주소'), D1: s('할인율'), E1: s('메모') };
for (let i = 0; i < 250; i++) {
  const r = i + 2;
  const special = ['Kim', 'kim', 'CON', '홍길동.', '홍길동'];
  mm[`A${r}`] = s(i < special.length ? special[i] : `회원${i}`);
  mm[`B${r}`] = n(46037 + i, 'yyyy-mm-dd');
  mm[`C${r}`] = s(`서울시 중구\n${i}번지 & <본관>`);
  mm[`D${r}`] = n(0.1, '0%');
  mm[`E${r}`] = s(i === 0 ? '' : '확인');
}
writeXlsx('mmdata.xlsx', [['회원', mm]]);

// DOCX template: variables in body + header + a docxtemplater loop tag that must not become an input field.
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
function docx(file, mainType) {
  const z = new PizZip();
  z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${mainType}.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`);
  z.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
  z.file('word/document.xml', `<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t xml:space="preserve">{{이</w:t></w:r><w:r><w:t>름}}님 가입일 {{가입일}}</w:t></w:r></w:p><w:p><w:r><w:t>주소: {{주소}} / 할인 {{할인율}}</w:t></w:r></w:p><w:p><w:r><w:t>{{#목록}}{{/목록}}</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr></w:body></w:document>`);
  z.file('word/header1.xml', `<?xml version="1.0"?><w:hdr ${W}><w:p><w:r><w:t>{{회사명}} 귀중</w:t></w:r></w:p></w:hdr>`);
  fs.writeFileSync(path.join(D, file), z.generate({ type: 'nodebuffer' }));
}
docx('letter.docx', 'document');
docx('letter.dotx', 'template');

// HWPX template with a run-split placeholder and the address variable (multi-line + XML special chars).
const sec = '<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"><hp:p><hp:run charPrIDRef="0"><hp:t>{{이</hp:t></hp:run><hp:run charPrIDRef="1"><hp:t>름}}님 / {{주소}}</hp:t></hp:run></hp:p></hs:sec>';
const hz = new PizZip();
hz.file('mimetype', 'application/hwp+zip');
hz.file('Contents/section0.xml', sec);
fs.writeFileSync(path.join(D, 'letter.hwpx'), hz.generate({ type: 'nodebuffer' }));

console.log('fixtures:', fs.readdirSync(D).filter((f) => !f.endsWith('.js')).join(' '));
