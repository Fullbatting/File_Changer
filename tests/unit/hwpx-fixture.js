// tests/unit/hwpx-fixture.js
// 테스트용 .hwpx 만들기: 표 칸 목록(hwp-reader.js의 cells와 같은 모양)을 한글이 저장하는 것과 같은 구조
// (hs:sec > hp:p > hp:run > hp:tbl > hp:tr > hp:tc > hp:subList > hp:p > hp:run > hp:t)로 쓴다.
// 실제 파일처럼 다양하게: 한 줄 글자를 런 두 개로 나누고(글자 모양이 바뀐 것처럼), 빈 칸은 세 가지 모양으로 둔다.
// 칸 글자의 한 줄이 "{RECT}"면 떠 있는 점선 상자를, "{TBL}"이면 문단 안에 들어앉은 작은 표를 넣는다(공통서식에 흔함).

const PizZip = require('pizzip');

const NS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const LINESEG = '<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="7000" flags="393216"/></hp:linesegarray>';

const P_OPEN = '<hp:p id="0" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">';
const RECT = `<hp:rect id="7" zOrder="1" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="8" ratio="0"><hp:offset x="0" y="0"/><hp:lineShape color="#000000" width="33" style="DOT"/><hp:drawText lastWidth="1000" name="" editable="0"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER">${P_OPEN}<hp:run charPrIDRef="3"/></hp:p></hp:subList></hp:drawText></hp:rect>`;

let emptyStyle = 0;
/** 문단 하나. 긴 줄은 가운데에서 런 두 개로 나눈다 */
function paragraph(line) {
  if (line === '{RECT}') return `${P_OPEN}<hp:run charPrIDRef="3">${RECT}</hp:run>${LINESEG}</hp:p>`;
  if (line === '{TBL}') return `${P_OPEN}<hp:run charPrIDRef="3">${table([{ row: 0, col: 0, text: '안쪽 표' }])}<hp:t/></hp:run>${LINESEG}</hp:p>`;
  if (line === '') {
    // 빈 칸의 세 가지 모양: 글자 없는 런 / 빈 <hp:t/> / 빈 <hp:t></hp:t>
    emptyStyle = (emptyStyle + 1) % 3;
    const run = emptyStyle === 0 ? '<hp:run charPrIDRef="3"/>' : emptyStyle === 1 ? '<hp:run charPrIDRef="3"><hp:t/></hp:run>' : '<hp:run charPrIDRef="3"><hp:t></hp:t></hp:run>';
    return `<hp:p id="0" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${run}${LINESEG}</hp:p>`;
  }
  const cut = line.length > 6 ? Math.floor(line.length / 2) : line.length;
  const runs = [line.slice(0, cut), line.slice(cut)].filter((x) => x !== '').map((x, i) => `<hp:run charPrIDRef="${4 + i}"><hp:t>${esc(x)}</hp:t></hp:run>`);
  return `<hp:p id="0" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">${runs.join('')}${LINESEG}</hp:p>`;
}

function table(cells) {
  const rowCnt = Math.max(...cells.map((c) => c.row + (c.rowSpan || 1)));
  const colCnt = Math.max(...cells.map((c) => c.col + (c.colSpan || 1)));
  const rows = [];
  for (let r = 0; r < rowCnt; r++) {
    const tcs = cells
      .filter((c) => c.row === r)
      .sort((a, b) => a.col - b.col)
      .map((c) => {
        const ps = String(c.text || '').split('\n').map(paragraph).join('');
        return `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3"><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${ps}</hp:subList><hp:cellAddr colAddr="${c.col}" rowAddr="${c.row}"/><hp:cellSpan colSpan="${c.colSpan || 1}" rowSpan="${c.rowSpan || 1}"/><hp:cellSz width="7000" height="1000"/><hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>`;
      });
    rows.push(`<hp:tr>${tcs.join('')}</hp:tr>`);
  }
  return `<hp:tbl id="1" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowCnt}" colCnt="${colCnt}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">${rows.join('')}</hp:tbl>`;
}

/** tables: [{ cells }], paragraphsAfter: 표 뒤 본문 문단 글자들. 반환: .hwpx 버퍼 */
function buildHwpx(tables, paragraphsAfter = []) {
  emptyStyle = 0;
  const body = tables.map((t) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${table(t.cells)}<hp:t/></hp:run>${LINESEG}</hp:p>`).join('');
  const after = paragraphsAfter.map(paragraph).join('');
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${body}${after}</hs:sec>`;
  const zip = new PizZip();
  zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
  zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1"/>');
  zip.file('Contents/section0.xml', xml);
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { buildHwpx };
