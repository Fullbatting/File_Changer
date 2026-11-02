const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const PizZip = require('pizzip');

const DIR = process.env.LEGACY_FIX;

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

