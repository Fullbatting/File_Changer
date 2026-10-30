// output-check.js
// ④ 변환 파일 확인: 저장한 엑셀 파일을 다시 열어(SheetJS) 시트·헤더·행 수·칸 값·표시 형식이
// 저장하려던 표와 같은지 비교한다. 화면에 기대지 않는 순수 함수라 node로 단위 테스트한다
// (tests/unit/output-check.test.js).

const XLSX = require('xlsx');

const MAX_LISTED_PROBLEMS = 50;
const DATE_TOLERANCE_DAYS = 1e-6; // 약 0.09초 — 초 단위까지 같은지 본다

const isEmpty = (v) => v === '' || v === null || v === undefined;

/** 화면에 보일 짧은 값 글자 */
function show(v) {
  if (isEmpty(v)) return '(빈 칸)';
  if (v instanceof Date) {
    const pad = (n) => String(n).padStart(2, '0');
    const base = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    return v.getHours() || v.getMinutes() || v.getSeconds() ? `${base} ${pad(v.getHours())}:${pad(v.getMinutes())}:${pad(v.getSeconds())}` : base;
  }
  const s = String(v);
  return s.length > 40 ? `${s.slice(0, 40)}…` : s;
}

const showCell = (cell) => (!cell || isEmpty(cell.v) ? '(빈 칸)' : `${show(cell.w !== undefined ? cell.w : cell.v)}${cell.t === 's' ? '(글자)' : cell.t === 'n' ? '(숫자)' : ''}`);

/**
 * 칸 하나 비교. 같으면 null, 다르면 { expected, actual, reason }.
 * format은 저장할 때 이 칸에 준 표시 형식(없으면 null) — 날짜·숫자는 형식까지 본다.
 */
function compareCell(raw, cell, format, toSerial) {
  if (isEmpty(raw)) {
    if (!cell || isEmpty(cell.v)) return null;
    return { expected: '(빈 칸)', actual: showCell(cell), reason: '빈 칸이어야 함' };
  }
  if (!cell || isEmpty(cell.v)) return { expected: show(raw), actual: '(빈 칸)', reason: '값이 없음' };
  if (raw instanceof Date) {
    if (cell.t !== 'n' || Math.abs(cell.v - toSerial(raw)) > DATE_TOLERANCE_DAYS) return { expected: show(raw), actual: showCell(cell), reason: '날짜가 다름' };
    if (format && cell.z !== format) return { expected: `${show(raw)} [${format}]`, actual: `[${cell.z || '형식 없음'}]`, reason: '날짜 형식이 다름' };
    return null;
  }
  if (typeof raw === 'number') {
    if (cell.t !== 'n' || Math.abs(cell.v - raw) > 1e-9 * Math.max(1, Math.abs(raw))) return { expected: `${show(raw)}(숫자)`, actual: showCell(cell), reason: '숫자가 다름' };
    if (format && cell.z !== format) return { expected: `${show(raw)} [${format}]`, actual: `[${cell.z || '형식 없음'}]`, reason: '숫자 형식이 다름' };
    return null;
  }
  if (typeof raw === 'boolean') {
    return cell.t === 'b' && cell.v === raw ? null : { expected: show(raw), actual: showCell(cell), reason: '참/거짓 값이 다름' };
  }
  const text = String(raw);
  // 글자는 글자로(00123이 숫자 123으로 바뀌지 않았는지까지) 본다
  if (cell.t !== 's' || cell.v !== text) return { expected: `${show(text)}(글자)`, actual: showCell(cell), reason: '글자가 다름' };
  return null;
}

/**
 * 다시 연 워크북을 저장하려던 내용과 비교한다.
 * spec: {
 *   headers: 값을 찾을 컬럼키, labels: 헤더 행 이름,
 *   sheets: [{ name, rows }] (저장한 순서),
 *   formatOf(columnIndex, raw): 저장할 때 준 표시 형식 또는 null,
 *   toSerial(date): 날짜 → 엑셀 일련번호
 * }
 * 반환: { ok, problemCount, problems: [{ kind, sheet, rowNumber, column, expected, actual, reason }],
 *        checkedCells, sheetNames, rowCounts: { 시트: 데이터 행 수 }, headerOk }
 */
function verifyWorkbook(workbook, { headers, labels, sheets, formatOf = () => null, toSerial }) {
  const problems = [];
  let problemCount = 0;
  let checkedCells = 0;
  let headerOk = true;
  const add = (p) => {
    problemCount += 1;
    if (problems.length < MAX_LISTED_PROBLEMS) problems.push(p);
  };

  const expectedNames = sheets.map((s) => s.name);
  if (JSON.stringify(workbook.SheetNames) !== JSON.stringify(expectedNames)) {
    add({ kind: 'sheets', expected: expectedNames.join(', '), actual: workbook.SheetNames.join(', '), reason: '시트 이름·순서가 다름' });
  }

  const rowCounts = {};
  sheets.forEach((s) => {
    const ws = workbook.Sheets[s.name];
    if (!ws) return;
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
    const lastRow = range ? range.e.r : -1;
    rowCounts[s.name] = Math.max(0, lastRow);

    labels.forEach((label, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      const actual = cell && !isEmpty(cell.v) ? String(cell.v) : '';
      if (actual !== String(label)) {
        headerOk = false;
        add({ kind: 'header', sheet: s.name, rowNumber: 1, column: label, expected: String(label), actual: actual || '(빈 칸)', reason: '헤더 이름이 다름' });
      }
    });
    if (range && range.e.c >= labels.length) {
      for (let r = 0; r <= range.e.r; r++) {
        for (let c = labels.length; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell && !isEmpty(cell.v)) {
            add({ kind: 'extra', sheet: s.name, rowNumber: r + 1, column: XLSX.utils.encode_col(c), expected: '(없음)', actual: showCell(cell), reason: '헤더 밖에 값이 있음' });
          }
        }
      }
    }

    // 끝쪽 행이 모두 빈 칸이면 파일에 행이 안 남을 수 있다 — 그 경우는 값 비교로 확인되므로 행 수 차이로 보지 않는다
    let lastFilled = s.rows.length - 1;
    while (lastFilled >= 0 && headers.every((h) => isEmpty(s.rows[lastFilled][h]))) lastFilled -= 1;
    if (lastRow > s.rows.length || lastRow < lastFilled + 1) {
      add({ kind: 'rows', sheet: s.name, expected: `${s.rows.length}행`, actual: `${Math.max(0, lastRow)}행`, reason: '행 수가 다름' });
    }

    s.rows.forEach((row, ri) => {
      headers.forEach((h, c) => {
        checkedCells += 1;
        const raw = row[h];
        const diff = compareCell(raw, ws[XLSX.utils.encode_cell({ r: ri + 1, c })], formatOf(c, raw), toSerial);
        if (diff) add({ kind: 'value', sheet: s.name, rowNumber: ri + 2, column: labels[c], ...diff });
      });
    });
  });

  return { ok: problemCount === 0, problemCount, problems, checkedCells, sheetNames: workbook.SheetNames.slice(), rowCounts, headerOk };
}

/** 문제 하나를 한 줄로 */
function describeProblem(p) {
  const where = [p.sheet ? `[${p.sheet}]` : '', p.rowNumber ? `${p.rowNumber}행` : '', p.column ? `"${p.column}"` : ''].filter(Boolean).join(' ');
  return `${where ? `${where}: ` : ''}${p.reason} — 저장하려던 값 ${p.expected} / 파일 ${p.actual}`;
}

module.exports = { verifyWorkbook, compareCell, describeProblem, MAX_LISTED_PROBLEMS };
