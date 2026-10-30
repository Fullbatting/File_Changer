// excel-writer.js
// [Load] 변환 결과를 서식이 적용된 엑셀(.xlsx)로 쓴다(exceljs). 화면에 기대지 않아 node로 단위 테스트한다.
// - 파일로 쓸 때는 exceljs 스트리밍 쓰기로 행을 차례로 내보내 큰 결과도 메모리에 한꺼번에 들고 있지 않고,
//   일정 행마다 진행률을 알리고 취소 요청을 확인한다.
// - formatOf(열 번호, 값)는 칸에 줄 표시 형식 규칙이다. ④ 변환 파일 확인(output-check.js)이 같은 규칙으로
//   다시 연 파일의 형식을 비교한다.

const ExcelJS = require('exceljs');

const isEmptyValue = (v) => v === '' || v === null || v === undefined;
const isNumericCell = (v) => typeof v === 'number' && Number.isFinite(v);
const PROGRESS_EVERY_ROWS = 2000;

/**
 * 문자열의 화면상 표시 폭을 추정한다. 한글/한자 등 전각 문자는 영문자보다
 * 넓게 표시되므로 가중치를 주어, 엑셀 열 너비를 실제 내용에 맞게 계산할 때
 * 한글이 많이 섞인 데이터에서도 "###" 잘림이 생기지 않도록 한다.
 */
function visualWidth(str) {
  let w = 0;
  for (const ch of String(str)) {
    const code = ch.codePointAt(0);
    const isWide =
      (code >= 0x1100 && code <= 0x11ff) || // 한글 자모
      (code >= 0x3130 && code <= 0x318f) || // 한글 호환 자모
      (code >= 0xac00 && code <= 0xd7a3) || // 한글 완성형 음절
      (code >= 0x4e00 && code <= 0x9fff); // CJK 통합 한자
    w += isWide ? 1.8 : 1;
  }
  return w;
}

/** Date를 엑셀 날짜 일련번호로. 시간대 계산 없이 화면에 보이는 연/월/일/시/분/초 그대로 옮긴다 */
function toExcelDateSerial(d) {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
  return (utc - Date.UTC(1899, 11, 30)) / 86400000;
}

/** 열 너비 계산용 글자(날짜는 화면 표기) */
function displayText(v) {
  if (v instanceof Date) {
    const pad = (n) => String(n).padStart(2, '0');
    const base = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    return v.getHours() || v.getMinutes() || v.getSeconds() ? `${base} 00:00:00` : base;
  }
  return String(v);
}

/**
 * 칸마다 줄 표시 형식 규칙을 만든다. 반환: (열 번호, 값) => 형식 글자 | null
 * - 날짜: "yyyy-mm-dd"(시각이 있으면 "yyyy-mm-dd hh:mm:ss").
 *   JS Date를 exceljs에 그대로 넘기면 UTC 기준으로 저장되어 한국 시간에서는 하루 전
 *   날짜가 되고 표시 형식도 미국식(mm-dd-yy)이 되므로 직접 일련번호로 바꾼다.
 * - 숫자 열(빈 값을 뺀 모든 값이 숫자, 또는 columnKinds가 number): 원본 형식(numberFormats) 또는
 *   천단위 구분(#,##0, 소수가 있으면 #,##0.######). 숫자 열 안에서 글자로 고친 칸은 형식 없이 글자로.
 *   숫자로 저장하는 것은 실제 숫자 값뿐이다 — 사번·계좌번호처럼 글자로 둔 값은 바뀌지 않는다.
 */
function makeCellFormatter(headers, rows, { numberFormats = {}, columnKinds = null } = {}) {
  const columnIsNumeric = columnKinds
    ? columnKinds.map((k) => k.kind === 'number')
    : headers.map((h) => {
        let sawAny = false;
        for (const row of rows) {
          const v = row[h];
          if (isEmptyValue(v)) continue;
          sawAny = true;
          if (!isNumericCell(v)) return false;
        }
        return sawAny;
      });
  const columnHasDecimal = headers.map((h, i) => columnIsNumeric[i] && rows.some((row) => isNumericCell(row[h]) && !Number.isInteger(row[h])));
  return (i, raw) => {
    if (raw instanceof Date) {
      const hasTime = raw.getHours() !== 0 || raw.getMinutes() !== 0 || raw.getSeconds() !== 0;
      return hasTime ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd';
    }
    if (columnIsNumeric[i] && isNumericCell(raw)) return numberFormats[headers[i]] || (columnHasDecimal[i] ? '#,##0.######' : '#,##0');
    return null;
  };
}

const HEADER_STYLE = {
  font: { bold: true, color: { argb: 'FFFFFFFF' } },
  fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } },
  alignment: { vertical: 'middle', horizontal: 'center' },
  border: {
    top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
  },
};

/** 시트 하나의 열 너비(헤더·값의 표시 폭, 8~40) */
function columnWidths(headers, labels, rows) {
  return headers.map((h, i) => {
    let maxWidth = visualWidth(labels[i]);
    for (const row of rows) {
      const v = row[h];
      if (!isEmptyValue(v)) maxWidth = Math.max(maxWidth, visualWidth(displayText(v)));
    }
    return Math.min(40, Math.max(8, Math.ceil(maxWidth) + 2));
  });
}

// 칸 스타일은 같은 객체를 함께 쓴다(칸마다 새 객체를 만들면 쓰기가 눈에 띄게 느려진다)
const TEXT_STYLE = Object.freeze({ alignment: Object.freeze({ horizontal: 'left' }) });
const styleCache = new Map();
function styleFor(fmt, isDate) {
  if (!fmt) return TEXT_STYLE;
  const key = `${isDate ? 'd' : 'n'}|${fmt}`;
  if (!styleCache.has(key)) styleCache.set(key, Object.freeze({ numFmt: fmt, alignment: Object.freeze({ horizontal: isDate ? 'center' : 'right' }) }));
  return styleCache.get(key);
}

class CancelledError extends Error {
  constructor() {
    super('사용자가 저장을 취소했습니다.');
    this.cancelled = true;
  }
}

/**
 * 워크북에 시트들을 쓴다(일반·스트리밍 공통). 스트리밍이면 행마다 commit해 메모리에서 내보낸다.
 * - 헤더 행: 배경색 + 흰 글자 + 굵게 + 테두리, 첫 행 고정(스크롤해도 헤더 유지)
 * - 열 너비: 헤더/값의 실제 표시 폭에 맞춰 자동 조정 → "###" 잘림 방지
 * onProgress(완료 행, 전체 행)는 PROGRESS_EVERY_ROWS행마다(그때 화면이 그려지도록 잠깐 양보한다),
 * shouldCancel()이 true면 CancelledError.
 */
async function writeSheets(workbook, streaming, { headers, labels, sheets, formatOf, onProgress, shouldCancel }) {
  const total = sheets.reduce((n, s) => n + s.rows.length, 0);
  let done = 0;
  for (const s of sheets) {
    const widths = columnWidths(headers, labels, s.rows);
    const worksheet = workbook.addWorksheet(s.name || 'Converted', { views: [{ state: 'frozen', ySplit: 1 }] });
    worksheet.columns = widths.map((width) => ({ width }));
    const headerRow = worksheet.addRow(labels);
    headerRow.eachCell((cell) => Object.assign(cell, HEADER_STYLE));
    if (streaming) headerRow.commit();

    for (const row of s.rows) {
      const values = headers.map((h) => {
        const raw = row[h];
        if (isEmptyValue(raw)) return '';
        if (raw instanceof Date) return toExcelDateSerial(raw);
        return raw;
      });
      const excelRow = worksheet.addRow(values);
      headers.forEach((h, i) => {
        const raw = row[h];
        if (isEmptyValue(raw)) return;
        excelRow.getCell(i + 1).style = styleFor(formatOf(i, raw), raw instanceof Date);
      });
      if (streaming) excelRow.commit();
      done += 1;
      if (done % PROGRESS_EVERY_ROWS === 0) {
        if (shouldCancel && shouldCancel()) throw new CancelledError();
        if (onProgress) onProgress(done, total);
        await new Promise((r) => setTimeout(r, 0));
      }
    }
    if (streaming) worksheet.commit();
  }
  if (onProgress) onProgress(total, total);
}

/**
 * 변환 결과를 엑셀로 쓴다.
 * spec: { headers, labels, sheets: [{ name, rows }], numberFormats, columnKinds, onProgress, shouldCancel }
 * filePath를 주면 그 파일로 스트리밍해 쓰고(반환 undefined), 없으면 Buffer를 돌려준다.
 */
async function writeStyledExcel(spec, filePath = null) {
  const allRows = spec.sheets.length === 1 ? spec.sheets[0].rows : spec.sheets.flatMap((s) => s.rows);
  const formatOf = spec.formatOf || makeCellFormatter(spec.headers, allRows, spec);
  const opts = { ...spec, formatOf };
  if (filePath) {
    // 공유 문자열 표 대신 칸 안에 글자를 넣는다(inlineStr): 값이 대부분 서로 다른 큰 표에서 쓰기·다시 읽기가
    // 빠르고 파일도 작다. 엑셀·리브레오피스 모두 읽는 표준 형식이다.
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: filePath, useStyles: true, useSharedStrings: false });
    await writeSheets(workbook, true, opts);
    await workbook.commit();
    return undefined;
  }
  const workbook = new ExcelJS.Workbook();
  await writeSheets(workbook, false, opts);
  return workbook.xlsx.writeBuffer();
}

module.exports = { writeStyledExcel, makeCellFormatter, toExcelDateSerial, visualWidth, CancelledError, PROGRESS_EVERY_ROWS };
