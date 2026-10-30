const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');

const DIR = process.env.LEGACY_FIX;

// --- Tab1 fixtures: A(source) has synonym/exact columns, B(target) has a title row above header ---
const wbA = XLSX.utils.book_new();
const wsA = XLSX.utils.aoa_to_sheet([
  ['이름', '생년월일', '전화번호', '이메일주소'],
  ['홍길동', '1990-01-01', '010-1111-2222', 'hong@example.com'],
  ['김영희', '1992-05-05', '010-3333-4444', 'kim@example.com'],
]);
XLSX.utils.book_append_sheet(wbA, wsA, 'Sheet1');
XLSX.writeFile(wbA, path.join(DIR, 'sourceA.xlsx'));

const wbB = XLSX.utils.book_new();
const wsB = XLSX.utils.aoa_to_sheet([
  ['2026년 명단'], // title row -> header auto-detect should skip this
  ['성명', '생일', '연락처', '이메일', '비고'],
]);
XLSX.utils.book_append_sheet(wbB, wsB, 'Sheet1');
XLSX.writeFile(wbB, path.join(DIR, 'targetB.xlsx'));

// --- Tab2 batch data fixture ---
const wbData = XLSX.utils.book_new();
const wsData = XLSX.utils.aoa_to_sheet([
  ['이름', '날짜'],
  ['홍길동', new Date(2026, 0, 15)],
  ['김영희', new Date(2026, 1, 20)],
]);
XLSX.utils.book_append_sheet(wbData, wsData, 'Sheet1');
XLSX.writeFile(wbData, path.join(DIR, 'mmData.xlsx'));

// --- Tab2 docx template fixture (minimal valid docx with {{이름}}, {{날짜}}) ---
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>안녕하세요 {{이름}}님, 날짜: {{날짜}}</w:t></w:r></w:p></w:body>
</w:document>`;
const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const zip = new PizZip();
zip.file('[Content_Types].xml', contentTypesXml);
zip.file('_rels/.rels', relsXml);
zip.file('word/document.xml', documentXml);
const docxBuffer = zip.generate({ type: 'nodebuffer' });
fs.writeFileSync(path.join(DIR, 'template.docx'), docxBuffer);

