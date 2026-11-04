// data-file.js
// 일괄 생성의 데이터 파일 읽기: 엑셀(.xlsx/.xls/.xlsm)·CSV, 그리고 한글 문서(.hwpx/.hwpt/.hwp) 속 표를
// 시트처럼(표 하나 = 시트 하나) 읽어 헤더와 행으로 만든다. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.
// - 한글 .hwpx 표는 서식 엔진(form-model.js)으로, 구형 .hwp 표는 hwp-reader.js로 읽는다.
// - 헤더 행 찾기·값 읽기 규칙은 table-reader.js.

const path = require('path');
const XLSX = require('xlsx');
const tableReader = require('./table-reader.js');
const { readHwpTables, cellsToRows } = require('./hwp-reader.js');
const { modelFromHwpx } = require('./form-model.js');
const { safeReadFileSync, createFileCache } = require('./file-utils.js');

// 표 형태 데이터 파일: xlsm은 xlsx와 구조가 같아 함께 읽히고, 한글 문서와 그 서식 파일(hwpt)은 문서 안의 표를 시트처럼 읽는다
const DATA_FILE_FILTERS = [
  { name: 'Excel / CSV / 한글 문서 속 표', extensions: ['xlsx', 'xls', 'xlsm', 'csv', 'hwpx', 'hwpt', 'hwp'] },
  { name: 'All Files', extensions: ['*'] },
];

/**
 * .hwpx 문서의 모든 표를 문서 순서대로 [{ rows }]로. 칸 글자는 문단을 줄바꿈으로 잇고(탭은 빈칸으로),
 * 칸 안에 든 중첩 표의 글자는 그 표의 것으로 따로 센다. 병합은 cellsToRows가 푼다.
 */
function readHwpxTables(buffer) {
  let model;
  try {
    model = modelFromHwpx(buffer);
  } catch (err) {
    if (/Contents\/section/.test(err.message)) throw err;
    throw new Error(`한글 문서 본문을 해석할 수 없습니다. 파일이 손상되었을 수 있습니다. (${err.message})`);
  }
  const textOf = (cell) =>
    cell.paragraphIds
      .map((id) => model.paragraphById.get(id).text.replace(/\t/g, ' '))
      .join('\n')
      .trim();
  return model.tables.map((table) => ({ rows: cellsToRows(table.cells.map((c) => ({ row: c.row, col: c.col, rowSpan: c.rowSpan, colSpan: c.colSpan, text: textOf(c) }))) }));
}

/**
 * 한글 문서에서 읽은 표들을 "표 1", "표 2"… 시트로 담은 SheetJS 워크북. 글자가 하나도 없는 표는 뺀다.
 * 반환: { workbook, note } — note는 찾은 표 요약(로그에 보일 글)
 */
function workbookFromTables(rawTables) {
  const tables = rawTables.filter(({ rows }) => rows.some((r) => r.some((v) => v !== ''))).map(({ rows }, i) => ({ name: `표 ${i + 1}`, rows }));
  if (tables.length === 0) throw new Error('한글 문서에서 표를 찾을 수 없습니다. 데이터가 표(행/열) 형태로 들어 있는지 확인하세요.');

  const workbook = XLSX.utils.book_new();
  tables.forEach(({ name, rows }) => {
    // 빈 칸은 셀을 만들지 않아야(null) 완전히 빈 행이 엑셀처럼 "빈 행"으로 취급되어 데이터에서 빠진다
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows.map((r) => r.map((v) => (v === '' ? null : v)))), name);
  });
  // 결재란·제목 상자 같은 작은 표가 데이터 표보다 먼저 나오는 경우가 많아, 행이 가장 많은 표를 기본으로 고른다
  workbook.preferredSheetName = tables.reduce((best, t) => (t.rows.length > best.rows.length ? t : best)).name;
  workbook.sheetUnitLabel = '표';

  const describe = ({ name, rows }) => {
    const firstFilled = rows.find((r) => r.some((v) => v !== '')) || [];
    const preview = firstFilled.filter((v) => v !== '').slice(0, 3).join(', ').replace(/\s+/g, ' ');
    return `${name}(${rows.length}행×${rows[0].length}열: ${preview.length > 30 ? `${preview.slice(0, 30)}…` : preview})`;
  };
  return { workbook, note: `한글 문서에서 표 ${tables.length}개를 찾았습니다: ${tables.map(describe).join(' / ')}` };
}

const workbookCache = createFileCache(8);

/**
 * 데이터 파일을 워크북으로 읽는다(같은 파일·인코딩은 다시 읽지 않는다).
 * - .hwpx/.hwpt/.hwp: 문서 안의 표 하나하나를 시트 하나로
 * - 그 밖에는 확장자가 아니라 파일 앞부분(내용)으로 판단: 글자 파일(CSV)은 인코딩을 자동 판별(또는 encoding으로
 *   직접 지정)해 문자열로, 그 외(xlsx·xls·HTML 표)는 SheetJS로 읽는다.
 * workbook.fileInfo에 내용 형식·인코딩·확장자와 다른지·1904 날짜 체계 여부를 담는다.
 * onNote(글): 파일을 새로 읽을 때 알릴 것(한글 문서에서 찾은 표 목록)
 */
function readWorkbook(filePath, { encoding = 'auto', onNote } = {}) {
  return workbookCache.get(
    filePath,
    () => {
      const ext = path.extname(filePath).toLowerCase();
      // fileInfo: 인식 결과 화면에 보일 파일 정보(실제 내용 형식, 글자 인코딩, 확장자와 다른지)
      const fileInfo = { ext, contentKind: null, mismatch: '', encoding: null, encodingAuto: true, brokenChars: 0, textSource: false };
      let workbook;
      if (ext === '.hwp' || ext === '.hwpx' || ext === '.hwpt') {
        const buffer = safeReadFileSync(filePath);
        const built = workbookFromTables(ext === '.hwp' ? readHwpTables(buffer) : readHwpxTables(buffer));
        workbook = built.workbook;
        fileInfo.textSource = true;
        if (onNote) onNote(built.note);
      } else {
        const buffer = safeReadFileSync(filePath);
        // 확장자가 아니라 실제 내용으로 읽는 방법을 정한다(CSV를 .xlsx로 이름만 바꾼 파일 등)
        fileInfo.contentKind = tableReader.sniffContentKind(buffer);
        fileInfo.mismatch = tableReader.describeContentMismatch(ext, fileInfo.contentKind);
        if (fileInfo.contentKind === 'text') {
          const decoded = tableReader.decodeTextBuffer(buffer, encoding);
          Object.assign(fileInfo, { encoding: decoded.encoding, encodingAuto: decoded.auto, brokenChars: decoded.brokenChars, textSource: true });
          // raw:true — CSV 글자를 숫자·날짜로 추측 변환하지 않는다(전화번호·우편번호의 앞자리 0, '2026-01-15'가 날짜 일련번호로 바뀌는 것 방지)
          workbook = XLSX.read(decoded.text, { type: 'string', raw: true, dense: true });
        } else {
          // cellNF: 셀의 표시 형식(z)을 남겨 날짜 셀·앞자리 0 채움 셀을 알아본다(날짜는 table-reader의 readCellValue가 푼다)
          workbook = XLSX.read(buffer, { type: 'buffer', cellNF: true, dense: true });
        }
      }
      fileInfo.date1904 = tableReader.isDate1904(workbook);
      workbook.fileInfo = fileInfo;
      return workbook;
    },
    `${filePath}|${encoding}`
  );
}

/** 시트 목록과 처음 고를 시트, 단위 이름("시트" 또는 한글 문서의 "표") */
function sheetChoices(workbook) {
  const names = workbook.SheetNames || [];
  return {
    names,
    unit: workbook.sheetUnitLabel || '시트',
    selected: names.includes(workbook.preferredSheetName) ? workbook.preferredSheetName : names[0] || null,
  };
}

/**
 * 시트에서 헤더 행을 찾고(또는 직접 고른 행을 쓰고) 그 아래 행들을 { 컬럼키: 값 } 객체로 읽는다.
 * 자세한 규칙은 table-reader.js의 buildTableFromSheet. 반환값의 report에 인식 결과(헤더 행, 건너뛴 행, 이름 없는 열,
 * 인코딩, 확장자와 다른 내용)가 들어 있다.
 * options: { displayText, headerRowIndex(0부터 · null이면 자동), encoding('auto'|'utf-8'|'cp949'), onNote }
 */
function readSheetTable(filePath, sheetName, { displayText = false, headerRowIndex = null, encoding = 'auto', onNote } = {}) {
  const workbook = readWorkbook(filePath, { encoding, onNote });
  const target = sheetName || workbook.SheetNames[0];
  if (!target || !workbook.Sheets[target]) throw new Error(`시트를 찾을 수 없습니다${sheetName ? ` (${sheetName})` : ''}.`);
  const fileInfo = workbook.fileInfo || {};
  const table = tableReader.buildTableFromSheet(workbook.Sheets[target], target, { displayText, date1904: !!fileInfo.date1904, headerRowIndex });
  table.fileInfo = fileInfo;
  table.textSource = !!fileInfo.textSource;
  return table;
}

module.exports = { DATA_FILE_FILTERS, readHwpxTables, workbookFromTables, readWorkbook, sheetChoices, readSheetTable };
