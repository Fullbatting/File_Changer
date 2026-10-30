// Builds a realistic HWPX document containing several tables, mirroring the OWPML structure Hancom writes.
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

function tc(row, col, text, { rowSpan = 1, colSpan = 1, inner = '' } = {}) {
  const run = text === null ? '<hp:run charPrIDRef="0"/>' : `<hp:run charPrIDRef="0"><hp:t>${text}</hp:t></hp:run>`;
  return (
    '<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3">' +
    '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">' +
    `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${run}${inner}` +
    '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="7000" flags="393216"/></hp:linesegarray></hp:p>' +
    '</hp:subList>' +
    `<hp:cellAddr colAddr="${col}" rowAddr="${row}"/><hp:cellSpan colSpan="${colSpan}" rowSpan="${rowSpan}"/>` +
    '<hp:cellSz width="7000" height="1000"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>'
  );
}

function tbl(rowCnt, colCnt, trs) {
  return (
    `<hp:tbl id="1" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowCnt}" colCnt="${colCnt}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">` +
    '<hp:sz width="28000" widthRelTo="ABSOLUTE" height="5000" heightRelTo="ABSOLUTE" protect="0"/>' +
    '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>' +
    '<hp:outMargin left="283" right="283" top="283" bottom="283"/><hp:inMargin left="510" right="510" top="141" bottom="141"/>' +
    trs.map((tr) => `<hp:tr>${tr.join('')}</hp:tr>`).join('') +
    '</hp:tbl>'
  );
}

const para = (inner) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${inner}</hp:p>`;
const tablePara = (t) => para(`<hp:run charPrIDRef="0">${t}<hp:t/></hp:run>`);

const approvalTable = tbl(2, 3, [
  [tc(0, 0, '담당'), tc(0, 1, '팀장'), tc(0, 2, '부장')],
  [tc(1, 0, null), tc(1, 1, null), tc(1, 2, null)],
]);
const emptyTable = tbl(1, 1, [[tc(0, 0, null)]]);
const dataTable = tbl(6, 4, [
  [tc(0, 0, '2026년 하반기 직원 명단', { colSpan: 4 })],
  [tc(1, 0, '이름', { rowSpan: 2 }), tc(1, 1, '연락처', { colSpan: 2 }), tc(1, 3, '부서', { rowSpan: 2 })],
  [tc(2, 1, '전화번호'), tc(2, 2, '이메일')],
  [tc(3, 0, '홍길동'), tc(3, 1, '010-1111-2222'), tc(3, 2, 'hong@example.com'), tc(3, 3, '영업팀', { rowSpan: 2 })],
  [tc(4, 0, '김영희'), tc(4, 1, '010-3333-4444'), tc(4, 2, 'kim&amp;co@example.com')],
  [tc(5, 0, '이철수'), tc(5, 1, '010-5555-6666'), tc(5, 2, 'lee@example.com'), tc(5, 3, '개발<hp:lineBreak/>2팀')],
]);
const nested = tbl(1, 1, [[tc(0, 0, '내부표')]]);
const outerWithNested = tbl(1, 2, [[tc(0, 0, '참고'), tc(0, 1, '메모', { inner: `<hp:run charPrIDRef="0">${nested}</hp:run>` })]]);

const sectionXml =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>' +
  '<hs:sec xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">' +
  para('<hp:run charPrIDRef="0"><hp:t>직원 명단 보고</hp:t></hp:run>') +
  tablePara(approvalTable) +
  tablePara(emptyTable) +
  tablePara(dataTable) +
  tablePara(outerWithNested) +
  '</hs:sec>';

const zip = new PizZip();
zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
zip.file('META-INF/container.xml', '<?xml version="1.0" encoding="UTF-8"?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>');
zip.file('Contents/content.hpf', '<?xml version="1.0" encoding="UTF-8"?><opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:manifest><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="section0"/></opf:spine></opf:package>');
zip.file('Contents/section0.xml', sectionXml);
fs.writeFileSync(path.join(process.env.LEGACY_FIX, 'staff.hwpx'), zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
fs.writeFileSync(path.join(process.env.LEGACY_FIX, 'legacy.hwp'), Buffer.from('HWP Document File'));
console.log('wrote staff.hwpx and legacy.hwp');
