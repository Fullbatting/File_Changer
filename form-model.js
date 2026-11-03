// form-model.js
// 서식 파일을 "표·칸·문단" 모델로 읽는다. 기입 위치 인식(form-recognize.js)과 내용 채우기(form-fill.js)가
// 같은 모델을 쓴다. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.
// - .hwpx: 문단마다 글자와 그 글자가 원문 XML의 어디에 있는지(pieces)까지 기억해 채울 수 있다.
// - .hwp: 표 안 글자만 읽는다(구형 바이너리는 쓰지 않으므로 인식·미리보기용). 채우려면 .hwpx로 바꿔야 한다.

const PizZip = require('pizzip');
const { parseXml, attr, childrenByLocal, walk, decodeEntities } = require('./hwpx-xml.js');

/** 문단 안에서 글자로 보이는 요소: 줄바꿈·탭은 한 글자로 센다(바꿀 수는 없는 자리) */
const VIRTUAL_CHARS = { lineBreak: '\n', tab: '\t' };

/** zip 안의 본문 파일 이름을 순서대로 */
function sectionNames(zip) {
  return Object.keys(zip.files)
    .map((name) => ({ name, m: /^Contents\/section(\d+)\.xml$/i.exec(name) }))
    .filter((x) => x.m)
    .sort((a, b) => Number(a.m[1]) - Number(b.m[1]))
    .map((x) => x.name);
}

/**
 * 문단(<hp:p>) 하나의 글자와 원문 위치. 문단 안에 든 표(<hp:tbl>)의 글자는 그 표의 문단으로 따로 센다.
 * pieces: [{ kind:'text', node, plainStart, plainEnd, decoded } | { kind:'virtual', plainStart, plainEnd }]
 */
function readParagraph(xml, pEl) {
  const pieces = [];
  const runs = [];
  let text = '';
  const visitT = (tEl) => {
    tEl.children.forEach((c) => {
      if (c.type === 'text') {
        const decoded = decodeEntities(xml.slice(c.start, c.end));
        pieces.push({ kind: 'text', node: c, tEl, plainStart: text.length, plainEnd: text.length + decoded.length, decoded });
        text += decoded;
      } else if (c.type === 'el') {
        const ch = VIRTUAL_CHARS[c.local] || '';
        if (ch) {
          pieces.push({ kind: 'virtual', plainStart: text.length, plainEnd: text.length + 1 });
          text += ch;
        }
      }
    });
  };
  walk(pEl, (el) => {
    if (el.local === 'tbl') return false;
    if (el.local === 'run') runs.push(el);
    if (el.local === 't') {
      visitT(el);
      return false;
    }
    return true;
  });
  return { text, pieces, runs };
}

/** .hwpx 파일(버퍼)을 모델로 */
function modelFromHwpx(buffer) {
  const zip = new PizZip(buffer);
  const names = sectionNames(zip);
  if (names.length === 0) throw new Error('올바른 .hwpx 파일이 아닙니다. (Contents/section*.xml 없음)');
  const model = { source: 'hwpx', sections: [], paragraphs: [], tables: [] };
  const paragraphByEl = new Map();
  names.forEach((name, si) => {
    const xml = zip.file(name).asText();
    const root = parseXml(xml);
    model.sections.push({ name, xml, root });
    let pIndex = 0;
    walk(root, (el) => {
      if (el.local === 'p') {
        const p = { id: `${si}.${pIndex++}`, section: si, el, ...readParagraph(xml, el), cell: null };
        model.paragraphs.push(p);
        paragraphByEl.set(el, p);
      }
      return true;
    });
    walk(root, (el) => {
      if (el.local !== 'tbl') return true;
      const table = { id: `t${model.tables.length + 1}`, section: si, rowCnt: Number(attr(el, 'rowCnt')) || 0, colCnt: Number(attr(el, 'colCnt')) || 0, cells: [] };
      model.tables.push(table);
      childrenByLocal(el, 'tr').forEach((tr) => {
        childrenByLocal(tr, 'tc').forEach((tc) => {
          const addr = childrenByLocal(tc, 'cellAddr')[0];
          const span = childrenByLocal(tc, 'cellSpan')[0];
          const cell = {
            row: addr ? Number(attr(addr, 'rowAddr')) : 0,
            col: addr ? Number(attr(addr, 'colAddr')) : 0,
            rowSpan: span ? Math.max(1, Number(attr(span, 'rowSpan')) || 1) : 1,
            colSpan: span ? Math.max(1, Number(attr(span, 'colSpan')) || 1) : 1,
            paragraphIds: [],
          };
          const subList = childrenByLocal(tc, 'subList')[0] || tc;
          childrenByLocal(subList, 'p').forEach((pEl) => {
            const p = paragraphByEl.get(pEl);
            if (!p) return;
            p.cell = { tableId: table.id, row: cell.row, col: cell.col, rowSpan: cell.rowSpan, colSpan: cell.colSpan };
            cell.paragraphIds.push(p.id);
          });
          table.cells.push(cell);
        });
      });
      return true; // 칸 안에 든 표도 따로 센다
    });
  });
  finishModel(model);
  return model;
}

/** hwp-reader.js의 표 목록([{ cells:[{row,col,rowSpan,colSpan,text}] }])을 모델로 — 읽기 전용 */
function modelFromHwpTables(tables) {
  const model = { source: 'hwp', readOnly: true, sections: [], paragraphs: [], tables: [] };
  let n = 0;
  tables.forEach((t) => {
    const table = { id: `t${model.tables.length + 1}`, section: 0, rowCnt: 0, colCnt: 0, cells: [] };
    model.tables.push(table);
    (t.cells || []).forEach((c) => {
      const cell = { row: c.row, col: c.col, rowSpan: c.rowSpan || 1, colSpan: c.colSpan || 1, paragraphIds: [] };
      table.rowCnt = Math.max(table.rowCnt, cell.row + cell.rowSpan);
      table.colCnt = Math.max(table.colCnt, cell.col + cell.colSpan);
      String(c.text || '').split('\n').forEach((line) => {
        const p = { id: `0.${n++}`, section: 0, text: line, pieces: null, runs: null, cell: { tableId: table.id, row: cell.row, col: cell.col, rowSpan: cell.rowSpan, colSpan: cell.colSpan } };
        model.paragraphs.push(p);
        cell.paragraphIds.push(p.id);
      });
      table.cells.push(cell);
    });
  });
  finishModel(model);
  return model;
}

/** 찾아보기 편의: id로 문단, 칸 글자, (표, 행, 열) 위치의 칸 */
function finishModel(model) {
  model.paragraphById = new Map(model.paragraphs.map((p) => [p.id, p]));
  model.tables.forEach((t) => {
    t.cells.forEach((c) => {
      c.text = c.paragraphIds.map((id) => model.paragraphById.get(id).text).join('\n');
    });
    t.cells.sort((a, b) => a.row - b.row || a.col - b.col);
  });
  model.tableById = new Map(model.tables.map((t) => [t.id, t]));
}

/** 표 table에서 (row, col) 자리를 덮는 칸 */
function cellAt(table, row, col) {
  return table.cells.find((c) => c.row <= row && row < c.row + c.rowSpan && c.col <= col && col < c.col + c.colSpan) || null;
}

module.exports = { modelFromHwpx, modelFromHwpTables, cellAt, sectionNames };
