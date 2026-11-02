// Fixtures for the full feature audit. Excel files are written as Excel stores cells (typed values,
// number formats, merges, used ranges not starting at A1), so the app's read path is exercised realistically.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const D = process.env.LEGACY_FIX;

const s = (v) => ({ t: 's', v });
const n = (v, z) => (z ? { t: 'n', v, z } : { t: 'n', v });
function writeXlsx(file, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(([name, ws]) => XLSX.utils.book_append_sheet(wb, ws, name));
  fs.writeFileSync(path.join(D, file), XLSX.write(wb, { type: 'buffer', bookType: file.endsWith('.xls') ? 'biff8' : 'xlsx' }));
}

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
