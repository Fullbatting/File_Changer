// table-reader.js
// 파일 내용 확인 · 글자 인코딩 · 셀 값 읽기 · 헤더 행 찾기 · 표 만들기.
// 화면(DOM)과 Electron에 기대지 않는 순수 함수만 두어 node로 단위 테스트한다(tests/unit/table-reader.test.js).
// 모든 판단은 정해진 규칙으로만 한다 — 같은 파일은 언제나 같은 결과가 나온다.

const XLSX = require('xlsx');
const iconv = require('iconv-lite');

/** 헤더 후보로 살펴볼 시트 맨 위 행 수 */
const HEADER_SCAN_ROWS = 10;

/* ---------- 파일 내용 확인 ---------- */

/**
 * 파일 앞부분으로 실제 형식을 알아본다(확장자와 다를 수 있다 — 다른 프로그램이 HTML 표를 .xls로
 * 내려주거나, CSV를 .xlsx로 이름만 바꾼 경우).
 * 'zip'(xlsx·hwpx) | 'ole'(xls·hwp) | 'html' | 'xml' | 'text' | 'binary'
 */
function sniffContentKind(buffer) {
  if (!buffer || buffer.length === 0) return 'text';
  if (buffer.length >= 4 && buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04) return 'zip';
  if (buffer.length >= 4 && buffer[0] === 0xd0 && buffer[1] === 0xcf && buffer[2] === 0x11 && buffer[3] === 0xe0) return 'ole';
  const head = buffer.slice(0, 4096);
  if (head.includes(0)) {
    // UTF-16 글자 파일은 0 바이트가 섞여 있다
    const bom16 = head.length >= 2 && ((head[0] === 0xff && head[1] === 0xfe) || (head[0] === 0xfe && head[1] === 0xff));
    return bom16 ? 'text' : 'binary';
  }
  const start = head.toString('latin1').replace(/^﻿|^\xEF\xBB\xBF/, '').trimStart().slice(0, 512).toLowerCase();
  if (start.startsWith('<?xml')) return /<html|<table/.test(start) ? 'html' : 'xml';
  if (start.startsWith('<')) return /<html|<table|<!doctype html|<meta|<head|<body/.test(start) ? 'html' : 'xml';
  return 'text';
}

const CONTENT_KIND_LABELS = {
  zip: '엑셀 2007 이상(.xlsx)·한글(.hwpx) 형식',
  ole: '옛 엑셀(.xls)·한글(.hwp) 형식',
  html: 'HTML 표',
  xml: 'XML 표(엑셀 2003 XML 등)',
  text: '글자 파일(CSV)',
  binary: '알 수 없는 형식',
};

const EXPECTED_KIND_BY_EXT = {
  '.xlsx': ['zip'],
  '.xlsm': ['zip'],
  '.xltx': ['zip'],
  '.xltm': ['zip'],
  '.xls': ['ole'],
  '.csv': ['text'],
  '.hwpx': ['zip'],
  '.hwpt': ['zip'],
  '.hwp': ['ole'],
};

/** 확장자와 실제 내용이 다르면 안내 문장, 같으면 '' */
function describeContentMismatch(ext, kind) {
  const expected = EXPECTED_KIND_BY_EXT[ext];
  if (!expected || expected.includes(kind)) return '';
  return `확장자는 ${ext}이지만 실제 내용은 ${CONTENT_KIND_LABELS[kind] || kind}입니다. 내용에 맞춰 읽었습니다 — 날짜·숫자 형식이 원래 프로그램과 다를 수 있으니 미리보기를 확인하세요.`;
}

/* ---------- 글자 인코딩 ---------- */

function isValidUTF8(buffer) {
  let i = 0;
  const len = buffer.length;
  while (i < len) {
    const byte = buffer[i];
    if (byte <= 0x7f) {
      i += 1;
      continue;
    }
    let extraBytes;
    if (byte >= 0xc2 && byte <= 0xdf) extraBytes = 1;
    else if (byte >= 0xe0 && byte <= 0xef) extraBytes = 2;
    else if (byte >= 0xf0 && byte <= 0xf4) extraBytes = 3;
    else return false;
    if (i + extraBytes >= len) return false;
    for (let j = 1; j <= extraBytes; j++) {
      if ((buffer[i + j] & 0xc0) !== 0x80) return false;
    }
    i += extraBytes + 1;
  }
  return true;
}

const ENCODING_LABELS = { 'utf-8-bom': 'UTF-8(BOM)', 'utf-8': 'UTF-8', cp949: 'CP949(한글 윈도우)', 'utf-16': 'UTF-16' };

/**
 * 글자 파일을 문자열로 푼다. encoding이 'auto'면:
 * BOM 있음 → UTF-8 / UTF-16, 올바른 UTF-8 → UTF-8, 그 외(한글 윈도우 엑셀의 기본 저장) → CP949.
 * 사용자가 'utf-8'·'cp949'를 직접 고를 수도 있다(자동 판단이 틀렸을 때).
 * 반환: { text, encoding, auto, brokenChars } — brokenChars는 풀 수 없어 �로 바뀐 글자 수.
 */
function decodeTextBuffer(buffer, encoding = 'auto') {
  const hasBOM = buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
  const bom16le = buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe;
  const bom16be = buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff;
  let used = encoding;
  let text;
  if (encoding === 'auto') {
    if (hasBOM) used = 'utf-8-bom';
    else if (bom16le || bom16be) used = 'utf-16';
    else used = isValidUTF8(buffer) ? 'utf-8' : 'cp949';
  }
  if (used === 'utf-8-bom' || (used === 'utf-8' && hasBOM)) {
    text = buffer.slice(3).toString('utf-8');
    used = 'utf-8-bom';
  } else if (used === 'utf-16') {
    text = iconv.decode(buffer, bom16be ? 'utf-16be' : 'utf-16le').replace(/^﻿/, '');
  } else if (used === 'utf-8') {
    text = buffer.toString('utf-8');
  } else {
    text = iconv.decode(buffer, 'cp949');
    used = 'cp949';
  }
  const brokenChars = (text.match(/�/g) || []).length;
  return { text, encoding: used, auto: encoding === 'auto', brokenChars };
}

/* ---------- 셀 값 ---------- */

/** 워크북이 1904 날짜 체계(맥용 엑셀의 옛 기본값)인지. 이 경우 같은 일련번호가 4년 1일 뒤 날짜다 */
function isDate1904(workbook) {
  return !!(workbook && workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);
}

/**
 * 셀 하나의 값을 읽는다.
 * - 날짜 셀은 SheetJS의 Date 자동 변환(cellDates) 대신 일련번호를 연/월/일/시/분/초로
 *   직접 풀어 "현지 시각"의 Date로 만든다. cellDates는 1899년 기준 시각을 현지
 *   시간대로 계산하는데, 서울처럼 그 무렵 표준시(LMT)가 지금과 달랐던 지역에서는
 *   몇십 초가 어긋나 날짜가 전날로 읽힌다(2026-01-15 → 2026-01-14 23:59:08).
 *   1904 날짜 체계 파일은 date1904를 넘겨야 4년 어긋나지 않는다.
 *   시각만 있는 셀(일련번호 1 미만)은 "HH:MM" 글자로 돌려준다.
 * - "00000"처럼 앞자리 0을 채워 보이게 한 숫자 셀(우편번호 등)은 보이는 글자 그대로.
 * - displayText면 숫자도 엑셀 화면에 보이는 글자(천 단위 구분·%·통화 기호) 그대로 —
 *   메일머지 문서에는 0.15가 아니라 "15%"가 들어가야 하기 때문이다.
 */
function readCellValue(cell, displayText, date1904 = false) {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  if (cell.t === 'e') return cell.w || '';
  if (cell.t === 'd' && cell.v instanceof Date) return cell.v;
  if (cell.t === 'n') {
    const format = typeof cell.z === 'string' ? cell.z : '';
    if (format && XLSX.SSF.is_date(format)) {
      const p = XLSX.SSF.parse_date_code(cell.v, { date1904 });
      if (p) {
        const pad = (n) => String(n).padStart(2, '0');
        if (cell.v < 1) return `${pad(p.H)}:${pad(p.M)}${p.S ? `:${pad(p.S)}` : ''}`;
        return new Date(p.y, p.m - 1, p.d, p.H, p.M, p.S);
      }
    }
    if (cell.w !== undefined && (displayText || /^0+$/.test(format))) return cell.w;
  }
  return cell.v;
}

/**
 * 셀 값을 화면·결과 파일에 보일 글자로 바꾼다. 날짜(Date)는 시·분·초가
 * 전부 0인(순수 날짜) 값은 "YYYY-MM-DD"로, 그 외에는 "YYYY-MM-DD HH:MM:SS"로
 * 표기한다. 탭1의 컬럼 매핑 미리보기와 탭2의 메일머지 변수 치환이 함께 쓴다.
 */
function formatPreviewValue(val) {
  if (val === '' || val === null || val === undefined) return '';
  if (val instanceof Date) {
    const pad = (n) => String(n).padStart(2, '0');
    const datePart = `${val.getFullYear()}-${pad(val.getMonth() + 1)}-${pad(val.getDate())}`;
    const hasTime = val.getHours() !== 0 || val.getMinutes() !== 0 || val.getSeconds() !== 0;
    return hasTime ? `${datePart} ${pad(val.getHours())}:${pad(val.getMinutes())}:${pad(val.getSeconds())}` : datePart;
  }
  return String(val);
}

/**
 * 시트를 2차원 배열로 읽는다. 행/열 번호는 시트의 사용 범위(!ref) 시작점 기준이다 —
 * 표가 A1이 아니라 B2 같은 곳에서 시작해도 헤더 찾기와 데이터 읽기가 같은 배열을
 * 쓰므로 어긋나지 않는다. 서식만 남은 먼 빈 칸 때문에 범위가 쓸데없이 커지지 않도록
 * 실제 값이 있는 마지막 행/열까지만 읽는다. 세로 병합 셀은 한글 문서 표와 같은 규칙
 * (cellsToRows 설명 참고)으로 아래 칸까지 값을 채운다.
 */
function sheetToGrid(sheet, displayText, date1904 = false) {
  if (!sheet || !sheet['!ref']) return [];
  const start = XLSX.utils.decode_range(sheet['!ref']).s;
  let lastRow = start.r - 1;
  let lastCol = start.c - 1;
  Object.keys(sheet).forEach((key) => {
    if (key[0] === '!') return;
    const v = sheet[key] && sheet[key].v;
    if (v === undefined || v === null || v === '') return;
    const addr = XLSX.utils.decode_cell(key);
    if (addr.r > lastRow) lastRow = addr.r;
    if (addr.c > lastCol) lastCol = addr.c;
  });

  const grid = [];
  for (let r = start.r; r <= lastRow; r++) {
    const row = [];
    for (let c = start.c; c <= lastCol; c++) row.push(readCellValue(sheet[XLSX.utils.encode_cell({ r, c })], displayText, date1904));
    grid.push(row);
  }

  (sheet['!merges'] || []).forEach(({ s, e }) => {
    const col = s.c - start.c;
    const top = grid[s.r - start.r];
    if (!top || col < 0 || col >= top.length) return;
    for (let r = s.r + 1; r <= e.r; r++) {
      const row = grid[r - start.r];
      if (row) row[col] = top[col];
    }
  });
  return grid;
}

/* ---------- 헤더 행 ---------- */

const isBlank = (v) => v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

/** 숫자·날짜·전화번호·금액처럼 보이는 글자는 컬럼 이름이 아니라 값이다 */
const VALUE_LIKE_TEXT = /^[\s\d.,:/\-+()%₩$원]+$/;

/** 컬럼 이름으로 쓸 만한 글자인지: 글자형이고, 값처럼 생기지 않았고, 문장처럼 길지 않음 */
function isHeaderLikeText(v) {
  if (typeof v !== 'string') return false;
  const s = v.trim();
  return s !== '' && s.length <= 40 && !VALUE_LIKE_TEXT.test(s);
}

/** 행 미리보기 글자(헤더 행 고르기 목록·인식 결과에 쓴다) */
function previewRow(row, maxLength = 48) {
  const text = row
    .filter((v) => !isBlank(v))
    .map((v) => formatPreviewValue(v).replace(/\s+/g, ' ').trim())
    .join(', ');
  return text.length > maxLength ? `${text.slice(0, maxLength)}…` : text;
}

/**
 * 시트 맨 위 행들을 헤더 후보로 평가한다. 행마다 { index, filled, textRatio, likely, preview }.
 * likely(헤더로 보임) 조건 — 모두 만족해야 한다:
 * - 채워진 칸이 위쪽 행들 중 가장 많이 채워진 행의 절반 이상(2칸 이상). 제목 한 줄·작성자 줄을 거른다.
 * - 채워진 칸의 80% 이상이 컬럼 이름 같은 글자(숫자·날짜·전화번호가 아님). 값이 든 데이터 행을 거른다.
 * - 같은 글자가 거의 반복되지 않음(헤더 이름은 대개 서로 다르다).
 * 데이터 행이 헤더보다 칸이 더 많이 채워져 있어도(머리글 칸 하나가 비어 있을 때) 헤더가 뽑힌다.
 */
function headerRowCandidates(grid, scanLimit = HEADER_SCAN_ROWS) {
  const n = Math.min(scanLimit, grid.length);
  const profiles = [];
  let maxFilled = 0;
  for (let i = 0; i < n; i++) {
    const row = grid[i];
    let filled = 0;
    let text = 0;
    let repeated = 0;
    const seen = new Set();
    row.forEach((v) => {
      if (isBlank(v)) return;
      filled += 1;
      if (isHeaderLikeText(v)) text += 1;
      const key = formatPreviewValue(v).trim();
      if (seen.has(key)) repeated += 1;
      seen.add(key);
    });
    profiles.push({ filled, text, repeated, width: row.length });
    maxFilled = Math.max(maxFilled, filled);
  }
  const need = maxFilled <= 1 ? 1 : Math.max(2, Math.ceil(maxFilled * 0.5));
  return profiles.map((p, index) => ({
    index,
    filled: p.filled,
    textRatio: p.filled ? p.text / p.filled : 0,
    likely: p.filled >= need && p.text / p.filled >= 0.8 && p.repeated <= p.filled * 0.2,
    preview: previewRow(grid[index]),
  }));
}

/** 예전 방식 점수(헤더로 보이는 행이 없을 때만 쓴다): 채워진 칸이 많을수록, 제목 한 칸짜리는 0 */
function legacyHeaderScore(row) {
  const nonEmpty = row.filter((c) => !isBlank(c));
  if (nonEmpty.length === 0) return -1;
  if (nonEmpty.length === 1 && row.length > 1) return 0;
  const avgLen = nonEmpty.reduce((sum, c) => sum + String(c).trim().length, 0) / nonEmpty.length;
  return avgLen > 20 ? nonEmpty.length * 0.5 : nonEmpty.length;
}

/**
 * 시트 맨 위 최대 10행에서 헤더 행(0부터)을 고른다. 헤더로 보이는 행 중 가장 위의 것.
 * 그런 행이 없으면(숫자로 된 컬럼 이름 등) 예전 방식대로 가장 많이 채워진 행.
 */
function detectHeaderRowIndex(grid) {
  const first = headerRowCandidates(grid).find((c) => c.likely);
  if (first) return first.index;
  const scanLimit = Math.min(HEADER_SCAN_ROWS, grid.length);
  let bestIndex = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < scanLimit; i++) {
    const score = legacyHeaderScore(grid[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestIndex;
}

/* ---------- 표 ---------- */

/**
 * 시트 하나를 표로 읽는다. 헤더와 데이터를 같은 2차원 배열(sheetToGrid)에서 열 번호로 짝지으므로,
 * 헤더 칸의 앞뒤 공백("이름 ")이나 같은 이름의 컬럼이 둘 있어도 값이 엉뚱한 컬럼으로 가지 않는다.
 * options: { displayText, date1904, headerRowIndex(직접 고른 헤더 행, 0부터 · 없으면 자동) }
 * 반환:
 * - headers: 컬럼키. 같은 이름이 또 나오면 "비고_1", "비고_2"처럼 구분한다.
 * - labels: 화면/결과 파일에 보일 원래 이름(공백만 정리).
 * - rows: 모든 칸이 빈 행은 뺀다. rowNumbers[i]: rows[i]의 시트 행 번호(1부터).
 * - report: 인식 결과 — 무엇을 건너뛰고 무엇을 바꿨는지(화면에 보여 주고 사용자가 고칠 수 있게).
 * 머리글 칸이 비어 있어도 그 열에 값이 있으면 버리지 않고 "(이름 없음) C열"로 넣는다.
 * 가로로 병합된 머리글(예: "주소"가 두 칸)은 뒤 칸을 "주소_1"로 이어 받는다.
 */
function buildTableFromSheet(sheet, sheetName, { displayText = false, date1904 = false, headerRowIndex: chosen = null } = {}) {
  const grid = sheetToGrid(sheet, displayText, date1904);
  if (grid.length === 0) throw new Error(`"${sheetName}" 시트에 데이터가 비어 있습니다.`);
  const start = XLSX.utils.decode_range(sheet['!ref']).s;

  const candidates = headerRowCandidates(grid);
  const autoIndex = detectHeaderRowIndex(grid);
  const manual = Number.isInteger(chosen) && chosen >= 0 && chosen < grid.length;
  const headerRowIndex = manual ? chosen : autoIndex;
  const headerRow = grid[headerRowIndex];

  // 머리글 행을 가로로 덮는 병합: 뒤 칸 → 첫 칸의 열 번호
  const mergedFrom = new Map();
  (sheet['!merges'] || []).forEach(({ s, e }) => {
    if (s.r - start.r !== headerRowIndex || e.c <= s.c) return;
    for (let c = s.c + 1; c <= e.c; c++) mergedFrom.set(c - start.c, s.c - start.c);
  });

  const hasDataBelow = (index) => {
    for (let r = headerRowIndex + 1; r < grid.length; r++) if (!isBlank(grid[r][index])) return true;
    return false;
  };

  const columns = [];
  const usedKeys = new Set();
  const unnamedColumns = [];
  const mergedHeaderColumns = [];
  const droppedEmptyColumns = [];
  headerRow.forEach((cell, index) => {
    const columnLetter = XLSX.utils.encode_col(start.c + index);
    let label = formatPreviewValue(cell).replace(/\s+/g, ' ').trim();
    if (!label && mergedFrom.has(index)) {
      label = formatPreviewValue(headerRow[mergedFrom.get(index)]).replace(/\s+/g, ' ').trim();
      if (label) mergedHeaderColumns.push(columnLetter);
    }
    if (!label) {
      if (!hasDataBelow(index)) {
        droppedEmptyColumns.push(columnLetter);
        return;
      }
      label = `(이름 없음) ${columnLetter}열`;
      unnamedColumns.push(label);
    }
    let key = label;
    for (let n = 1; usedKeys.has(key); n++) key = `${label}_${n}`;
    usedKeys.add(key);
    columns.push({ key, label, index });
  });
  if (columns.length === 0) throw new Error(`"${sheetName}" 시트에서 헤더(컬럼명) 행을 찾을 수 없습니다.`);

  const rows = [];
  const rowNumbers = [];
  let emptyRowsRemoved = 0;
  for (let i = headerRowIndex + 1; i < grid.length; i++) {
    const r = grid[i];
    const obj = {};
    let any = false;
    columns.forEach(({ key, index }) => {
      const v = r[index] === undefined ? '' : r[index];
      obj[key] = v;
      if (v !== '') any = true;
    });
    if (!any) {
      emptyRowsRemoved += 1;
      continue;
    }
    rows.push(obj);
    rowNumbers.push(start.r + i + 1);
  }

  // 컬럼별 원본 숫자 표시 형식(%, 통화, 천 단위 등). 변환 결과 엑셀에서 0.15가 "15%"로 보이도록 넘겨준다.
  const numberFormats = {};
  columns.forEach(({ key, index }) => {
    for (let r = headerRowIndex + 1; r < grid.length; r++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: start.r + r, c: start.c + index })];
      if (cell && cell.t === 'n' && typeof cell.z === 'string' && cell.z !== 'General' && !XLSX.SSF.is_date(cell.z)) {
        numberFormats[key] = cell.z;
        break;
      }
    }
  });

  const skippedRows = [];
  for (let i = 0; i < headerRowIndex; i++) {
    if (grid[i].some((v) => !isBlank(v))) skippedRows.push({ rowNumber: start.r + i + 1, preview: previewRow(grid[i]) });
  }

  const report = {
    sheetName,
    headerRowIndex,
    headerRowNumber: start.r + headerRowIndex + 1,
    headerAuto: !manual,
    autoHeaderRowNumber: start.r + autoIndex + 1,
    headerCandidates: candidates.map((c) => ({ ...c, rowNumber: start.r + c.index + 1 })),
    firstRowNumber: start.r + 1,
    skippedRows,
    emptyRowsRemoved,
    unnamedColumns,
    mergedHeaderColumns,
    droppedEmptyColumns,
    date1904,
    dataRowCount: rows.length,
  };

  return {
    headers: columns.map((c) => c.key),
    labels: columns.map((c) => c.label),
    headerRowIndex,
    rows,
    rowNumbers,
    numberFormats,
    report,
  };
}

/**
 * 인식 결과(report + 파일 정보)를 사람이 읽을 짧은 문장 목록으로. 주의가 필요한 것은 warn: true.
 * fileInfo: { encoding, encodingAuto, brokenChars, mismatch }
 */
function describeRecognition(report, fileInfo = {}, { includeRows = true } = {}) {
  const items = [];
  items.push({
    text: `헤더: ${report.headerRowNumber}행${report.headerAuto ? '(자동)' : '(직접 고름)'}${includeRows ? ` · 데이터 ${report.dataRowCount}행` : ''}`,
    warn: false,
  });
  if (report.skippedRows.length > 0) {
    const list = report.skippedRows.slice(0, 3).map((r) => `${r.rowNumber}행 "${r.preview}"`).join(', ');
    items.push({ text: `헤더 위 ${report.skippedRows.length}개 행은 데이터에서 뺌: ${list}${report.skippedRows.length > 3 ? ' …' : ''}`, warn: true });
  }
  if (report.unnamedColumns.length > 0) {
    items.push({ text: `머리글이 빈 열 ${report.unnamedColumns.length}개를 ${report.unnamedColumns.join(', ')}(으)로 넣음(값이 있어 버리지 않음)`, warn: true });
  }
  if (report.mergedHeaderColumns.length > 0) {
    items.push({ text: `병합된 머리글을 이어 받은 열: ${report.mergedHeaderColumns.join(', ')}`, warn: false });
  }
  if (includeRows && report.emptyRowsRemoved > 0) items.push({ text: `완전히 빈 행 ${report.emptyRowsRemoved}개는 뺌`, warn: false });
  if (report.date1904) items.push({ text: '1904 날짜 체계 파일 — 날짜를 그 기준으로 읽음', warn: false });
  if (fileInfo.encoding) {
    const broken = fileInfo.brokenChars ? ` · 깨진 글자 ${fileInfo.brokenChars}개 — 인코딩을 바꿔 보세요` : '';
    items.push({ text: `글자 인코딩: ${ENCODING_LABELS[fileInfo.encoding] || fileInfo.encoding}${fileInfo.encodingAuto ? '(자동)' : '(직접 고름)'}${broken}`, warn: !!fileInfo.brokenChars });
  }
  if (fileInfo.mismatch) items.push({ text: fileInfo.mismatch, warn: true });
  return items;
}

module.exports = {
  HEADER_SCAN_ROWS,
  sniffContentKind,
  describeContentMismatch,
  isValidUTF8,
  decodeTextBuffer,
  ENCODING_LABELS,
  isDate1904,
  readCellValue,
  formatPreviewValue,
  sheetToGrid,
  isHeaderLikeText,
  headerRowCandidates,
  detectHeaderRowIndex,
  buildTableFromSheet,
  describeRecognition,
};
