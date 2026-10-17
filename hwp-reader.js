// hwp-reader.js
// 구형 한글(.hwp, HWP 5.0 바이너리) 문서에서 표를 읽어 2차원 배열로 돌려준다.
// DOM이 필요 없는 순수 Node 모듈이라 렌더러(common.js)와 단위 테스트에서 함께 쓴다.
//
// HWP 5.0 구조: OLE 복합 문서 안에 FileHeader(서명/버전/플래그)와 본문
// BodyText/Section0, Section1...이 들어 있고, 본문은 (보통) raw deflate로 압축된
// "레코드"의 나열이다. 각 레코드 헤더(4바이트)는 태그(10비트)·레벨(10비트)·
// 크기(12비트, 0xFFF이면 뒤의 4바이트가 실제 크기)로 이루어지며, 레벨로 트리를
// 이룬다. 표는 CTRL_HEADER('tbl ') 아래에 TABLE 레코드가 오고, 그 뒤로 셀마다
// LIST_HEADER(셀 주소·병합 정보)와 그 셀의 문단(PARA_HEADER > PARA_TEXT)들이
// 같은 레벨의 형제로 이어진다. TABLE보다 앞에 오는 LIST_HEADER는 표 캡션이다.
// 레이아웃은 한컴 공개 문서와, 실제 파일로 검증된 pyhwp 구현을 참고했다.
// (셀 LIST_HEADER는 공개 문서 표기(6바이트)와 달리 실제로는 8바이트 뒤에 셀
// 주소가 온다 — pyhwp의 정의와 실제 한컴 파일에서 확인.)

const zlib = require('zlib');
const XLSX = require('xlsx'); // SheetJS에 내장된 OLE 복합 문서(CFB) 리더를 쓴다

const TAG_PARA_HEADER = 66;
const TAG_PARA_TEXT = 67;
const TAG_CTRL_HEADER = 71;
const TAG_LIST_HEADER = 72;
const TAG_TABLE = 77;
const CTRL_ID_TABLE = 0x74626c20; // 't','b','l',' '

const OLE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const FLAG_COMPRESSED = 1 << 0;
const FLAG_PASSWORD = 1 << 1;
const FLAG_DISTRIBUTABLE = 1 << 2;

// 문단 글자 속 제어 문자 중 1글자(2바이트)짜리. 나머지 0~31은 8글자(16바이트) 크기다.
const SINGLE_WIDTH_CONTROLS = new Set([0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31]);

function hwpError(message) {
  return new Error(message);
}

/**
 * 셀 목록 [{ row, col, rowSpan, text }]을 2차원 배열로 바꾼다. 병합 셀 값은 왼쪽 위
 * 칸에 두고, 세로 병합(rowSpan)은 아래 칸에도 같은 값을 채운다 — 한글 표는
 * "부서"처럼 여러 행이 공유하는 값을 세로 병합으로 자주 표현하는데, 비워 두면
 * 둘째 행부터 값이 사라지고, 2단 헤더(위: "이름"(세로병합) | "연락처"(가로병합),
 * 아래: "전화번호" | "이메일")에서도 아래 행에 "이름"이 채워져야 그 행이 완전한
 * 헤더 행이 되기 때문이다. 가로 병합은 같은 이름의 컬럼이 여러 개 생기지 않도록
 * 채우지 않는다. HWPX(common.js)와 HWP가 함께 쓴다.
 */
function cellsToRows(cells) {
  let rowCount = 0;
  let colCount = 0;
  cells.forEach(({ row, col, rowSpan, colSpan }) => {
    rowCount = Math.max(rowCount, row + rowSpan);
    colCount = Math.max(colCount, col + (colSpan || 1));
  });
  const rows = Array.from({ length: rowCount }, () => Array(colCount).fill(''));
  cells.forEach(({ row, col, rowSpan, text }) => {
    for (let r = row; r < row + rowSpan; r++) rows[r][col] = text;
  });
  return rows;
}

/** 레코드를 읽어 레벨 기준 트리로 만든다 */
function parseRecordTree(data) {
  const root = { level: -1, children: [] };
  const stack = [root];
  let offset = 0;
  while (offset + 4 <= data.length) {
    const header = data.readUInt32LE(offset);
    offset += 4;
    const tag = header & 0x3ff;
    const level = (header >>> 10) & 0x3ff;
    let size = (header >>> 20) & 0xfff;
    if (size === 0xfff) {
      if (offset + 4 > data.length) break;
      size = data.readUInt32LE(offset);
      offset += 4;
    }
    if (offset + size > data.length) {
      throw hwpError('한글 문서 본문이 중간에 잘려 있습니다. 파일이 손상되었을 수 있습니다.');
    }
    const record = { tag, level, payload: data.subarray(offset, offset + size), children: [] };
    offset += size;

    while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop();
    stack[stack.length - 1].children.push(record);
    stack.push(record);
  }
  return root;
}

/** PARA_TEXT(UTF-16LE)를 글자로 바꾼다. 줄바꿈은 살리고, 탭/특수 공백은 공백으로, 나머지 제어 문자는 뺀다 */
function decodeParaText(payload) {
  const units = [];
  for (let i = 0; i + 1 < payload.length; ) {
    const code = payload.readUInt16LE(i);
    if (code >= 32) {
      units.push(code);
      i += 2;
      continue;
    }
    if (code === 10) units.push(10);
    else if (code === 9 || code === 30 || code === 31) units.push(32);
    else if (code === 24) units.push(45);
    i += SINGLE_WIDTH_CONTROLS.has(code) ? 2 : 16;
  }
  let text = '';
  for (let i = 0; i < units.length; i += 4096) {
    text += String.fromCharCode.apply(null, units.slice(i, i + 4096));
  }
  return text;
}

function paragraphText(paraNode) {
  const textRecord = paraNode.children.find((c) => c.tag === TAG_PARA_TEXT);
  return textRecord ? decodeParaText(textRecord.payload) : '';
}

function isTableControl(node) {
  return node.tag === TAG_CTRL_HEADER && node.payload.length >= 4 && node.payload.readUInt32LE(0) === CTRL_ID_TABLE;
}

/**
 * 표 컨트롤 하나의 셀을 읽는다. 셀 안의 문단에 다시 표가 들어 있으면(중첩 표)
 * 그 표는 tables에 따로 추가하고, 이 셀의 글자에는 넣지 않는다.
 */
function readTableCells(tableNode, tables) {
  const tableRecord = tableNode.children.find((c) => c.tag === TAG_TABLE);
  if (!tableRecord || tableRecord.payload.length < 8) return [];
  const rowLimit = tableRecord.payload.readUInt16LE(4);
  const colLimit = tableRecord.payload.readUInt16LE(6);

  const cells = [];
  let seenTableRecord = false;
  let current = null;
  tableNode.children.forEach((child) => {
    if (child === tableRecord) {
      seenTableRecord = true;
      return;
    }
    if (!seenTableRecord) return; // TABLE 레코드 앞의 LIST_HEADER/문단은 표 캡션
    if (child.tag === TAG_LIST_HEADER) {
      current = null;
      if (child.payload.length < 16) return;
      const col = child.payload.readUInt16LE(8);
      const row = child.payload.readUInt16LE(10);
      if (row >= rowLimit || col >= colLimit) return;
      current = {
        col,
        row,
        colSpan: Math.min(Math.max(1, child.payload.readUInt16LE(12)), colLimit - col),
        rowSpan: Math.min(Math.max(1, child.payload.readUInt16LE(14)), rowLimit - row),
        paragraphs: [],
      };
      cells.push(current);
    } else if (child.tag === TAG_PARA_HEADER && current) {
      current.paragraphs.push(paragraphText(child));
      collectTables(child, tables);
    }
  });

  return cells.map(({ paragraphs, ...cell }) => ({ ...cell, text: paragraphs.join('\n').trim() }));
}

/** 트리 전체에서 표를 문서에 나오는 순서대로(바깥 표가 그 안의 중첩 표보다 먼저) 모은다 */
function collectTables(node, tables) {
  node.children.forEach((child) => {
    if (isTableControl(child)) {
      const entry = { rows: [] };
      tables.push(entry);
      entry.rows = cellsToRows(readTableCells(child, tables));
    } else {
      collectTables(child, tables);
    }
  });
}

function readStream(cfb, fullPath) {
  const idx = cfb.FullPaths.indexOf(fullPath);
  return idx >= 0 && cfb.FileIndex[idx].content ? Buffer.from(cfb.FileIndex[idx].content) : null;
}

/** .hwp 파일 버퍼에서 모든 표를 문서 순서대로 [{ rows }]로 읽는다 (빈 표 포함) */
function readHwpTables(buffer) {
  if (buffer.slice(0, 20).toString('latin1') === 'HWP Document File V3') {
    throw hwpError('한글 97 이전(HWP 3.0) 형식의 문서는 지원하지 않습니다. 최신 한글 프로그램에서 다시 저장한 뒤 시도하세요.');
  }
  if (buffer.length < 512 || !buffer.subarray(0, 8).equals(OLE_SIGNATURE)) {
    throw hwpError('올바른 한글(.hwp) 문서가 아닙니다.');
  }

  let cfb;
  try {
    cfb = XLSX.CFB.read(buffer, { type: 'buffer' });
  } catch (err) {
    throw hwpError('한글(.hwp) 문서를 열 수 없습니다. 파일이 손상되었을 수 있습니다.');
  }

  const root = cfb.FullPaths[0];
  const fileHeader = readStream(cfb, `${root}FileHeader`);
  if (!fileHeader || fileHeader.length < 40 || fileHeader.slice(0, 17).toString('latin1') !== 'HWP Document File') {
    throw hwpError('올바른 한글(.hwp) 문서가 아닙니다.');
  }
  if (fileHeader[35] !== 5) {
    throw hwpError(`지원하지 않는 한글 문서 버전입니다 (${fileHeader[35]}.${fileHeader[34]}). 한글 2002 이후 형식(5.0)만 읽을 수 있습니다.`);
  }
  const flags = fileHeader.readUInt32LE(36);
  if (flags & FLAG_PASSWORD) {
    throw hwpError('암호가 걸린 한글 문서는 읽을 수 없습니다. 한글에서 암호를 해제한 뒤 다시 시도하세요.');
  }
  if (flags & FLAG_DISTRIBUTABLE) {
    throw hwpError('배포용으로 저장된 한글 문서는 내용이 암호화되어 있어 읽을 수 없습니다. 원본 문서를 사용하세요.');
  }

  const sectionPaths = cfb.FullPaths.filter((p) => /\/BodyText\/Section\d+$/.test(p)).sort(
    (a, b) => parseInt(a.match(/Section(\d+)$/)[1], 10) - parseInt(b.match(/Section(\d+)$/)[1], 10)
  );
  if (sectionPaths.length === 0) {
    throw hwpError('한글 문서에서 본문을 찾을 수 없습니다. 파일이 손상되었을 수 있습니다.');
  }

  const tables = [];
  sectionPaths.forEach((p) => {
    const raw = readStream(cfb, p) || Buffer.alloc(0);
    let data = raw;
    if (flags & FLAG_COMPRESSED) {
      try {
        data = zlib.inflateRawSync(raw);
      } catch (err) {
        throw hwpError('한글 문서 본문의 압축을 풀 수 없습니다. 파일이 손상되었을 수 있습니다.');
      }
    }
    collectTables(parseRecordTree(data), tables);
  });
  return tables;
}

module.exports = { readHwpTables, cellsToRows, decodeParaText };
