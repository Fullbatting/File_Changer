// common.js
// 메일머지 자동기입(tab2.js)이 쓰는 공통 유틸·캐시·파일 읽기(데이터 파일·한글 문서 표)·
// 동의어 사전·변수 자동 매핑 연결부.

const { ipcRenderer, webUtils } = require('electron');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const { readHwpTables, cellsToRows } = require('./hwp-reader.js');
const tableReader = require('./table-reader.js');
const autoMatch = require('./auto-match.js');
const { formatPreviewValue } = tableReader;

/* =========================================================================
   0. 공통 유틸 (로그, 탭 전환)
   ========================================================================= */

const logArea = document.getElementById('logArea');

let logScrollPending = false;

function log(message, level = 'info') {
  const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
  const prefix = { info: 'INFO', ok: ' OK ', warn: 'WARN', error: 'FAIL' }[level] || 'INFO';
  // value +=는 로그 전체를 다시 배치하므로 끝에 덧붙이기만 한다
  const end = logArea.value.length;
  logArea.setRangeText(`[${time}] [${prefix}] ${message}\n`, end, end, 'end');
  // 맨 아래로 스크롤은 다음 화면 갱신 때 한 번만(바로 scrollHeight를 읽으면 강제 레이아웃이 생긴다)
  if (logScrollPending) return;
  logScrollPending = true;
  requestAnimationFrame(() => {
    logScrollPending = false;
    logArea.scrollTop = logArea.scrollHeight;
  });
}

function notifyError(context, err) {
  const msg = friendlyErrorMessage(err);
  log(`${context}: ${msg}`, 'error');
  alert(`❌ ${context}\n\n${msg}`);
}

/** Node.js 파일 시스템 에러 코드를 이해하기 쉬운 한국어 메시지로 변환 */
function friendlyErrorMessage(err) {
  if (!err) return '알 수 없는 오류입니다.';
  const code = err.code;
  const target = err.path ? ` (${err.path})` : '';
  switch (code) {
    case 'ENOENT':
      return `파일을 찾을 수 없습니다${target}. 파일이 이동되었거나 삭제되지 않았는지 확인하세요.`;
    case 'EACCES':
    case 'EPERM':
      return `파일에 접근할 권한이 없습니다${target}. 관리자 권한 또는 파일 권한을 확인하세요.`;
    case 'EBUSY':
      return `파일이 다른 프로그램(Excel/Word 등)에서 열려 있어 사용할 수 없습니다${target}. 해당 프로그램에서 파일을 닫은 뒤 다시 시도하세요.`;
    case 'EISDIR':
      return '선택한 경로는 폴더입니다. 파일을 선택하세요.';
    case 'ENOSPC':
      return '디스크 여유 공간이 부족합니다.';
    default:
      return err.message ? String(err.message) : String(err);
  }
}

/** 파일을 안전하게 읽고, 실패 시 친절한 에러 메시지로 재발생시킴 */
function safeReadFileSync(filePath, encoding) {
  if (!fs.existsSync(filePath)) {
    const err = new Error(`파일을 찾을 수 없습니다 (${filePath})`);
    err.code = 'ENOENT';
    err.path = filePath;
    throw err;
  }
  try {
    return encoding ? fs.readFileSync(filePath, encoding) : fs.readFileSync(filePath);
  } catch (err) {
    err.path = err.path || filePath;
    throw err;
  }
}

/** 파일을 안전하게 저장하고, 실패 시 친절한 에러 메시지로 재발생시킴 */
function safeWriteFileSync(filePath, data) {
  try {
    fs.writeFileSync(filePath, data);
  } catch (err) {
    err.path = err.path || filePath;
    throw err;
  }
}

/** 파일의 현재 수정 시각(ms). 캐시 무효화 판단용. 파일이 없으면 null. */
function getFileMtimeMs(filePath) {
  return fs.existsSync(filePath) ? fs.statSync(filePath).mtimeMs : null;
}

/** Map 기반 캐시가 limit을 넘으면 가장 오래된(먼저 삽입된) 항목을 제거(단순 FIFO) */
function evictOldestIfFull(map, limit) {
  if (map.size >= limit) {
    map.delete(map.keys().next().value);
  }
}

/**
 * "진행 / 취소" 형태의 확인 모달 공통 열기/닫기 로직. 모달을 보이고, 진행·취소
 * 버튼 클릭 또는 배경(오버레이) 클릭(취소로 처리)에 따라 true/false를 Promise로
 * 반환한다. 모달에 채울 내용(요약/미리보기 테이블)을 만드는 로직은 호출하는 쪽에 둔다.
 */
function confirmModal(modalEl, proceedBtn, cancelBtn) {
  return new Promise((resolve) => {
    modalEl.hidden = false;

    const cleanup = () => {
      modalEl.hidden = true;
      proceedBtn.onclick = null;
      cancelBtn.onclick = null;
      modalEl.onclick = null;
    };

    proceedBtn.onclick = () => {
      cleanup();
      resolve(true);
    };
    cancelBtn.onclick = () => {
      cleanup();
      resolve(false);
    };
    modalEl.onclick = (e) => {
      if (e.target === modalEl) cancelBtn.click();
    };
  });
}

/* -------------------------------------------------------------------------
 * 한글 문서(HWPX/HWP)의 표를 데이터로 읽기
 * HWPX 본문(Contents/section*.xml)의 표(<hp:tbl> > <hp:tr> > <hp:tc>)를 2차원
 * 배열로 풀어(구형 .hwp는 hwp-reader.js가 같은 형태로 읽는다), 표 하나를 엑셀
 * 시트 하나처럼 다룰 수 있게 한다. SheetJS 워크북
 * 형태로 맞춰 두면 시트 선택·헤더 자동 감지·매핑·변환·일괄 생성 로직을
 * 엑셀/CSV와 완전히 똑같이 재사용할 수 있다. 표 안에 표가 또 들어 있는(중첩)
 * 경우를 정확히 가르기 위해 정규식 대신 DOMParser로 XML 구조를 그대로 따라간다.
 * ------------------------------------------------------------------------- */

/** hwpx zip에서 본문 섹션 XML 파일 이름 목록을 번호순으로 반환 */
function getHwpxSectionNames(zip) {
  return Object.keys(zip.files)
    .filter((name) => /^Contents\/section\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/section(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/section(\d+)\.xml/)[1], 10);
      return na - nb;
    });
}

function childElementsByLocalName(el, localName) {
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === localName);
}

function intAttribute(el, name, fallback) {
  const v = el ? parseInt(el.getAttribute(name), 10) : NaN;
  return Number.isNaN(v) ? fallback : v;
}

/** node를 감싸는 가장 가까운 <tbl>. 셀 안의 글자가 이 표의 것인지, 셀 속 중첩 표의 것인지 가르는 데 쓴다 */
function closestHwpxTable(node) {
  for (let cur = node.parentNode; cur && cur.nodeType === 1; cur = cur.parentNode) {
    if (cur.localName === 'tbl') return cur;
  }
  return null;
}

/** <hp:t> 하나의 글자. 그 안의 줄바꿈/탭/특수 공백 요소는 실제 문자로 바꾼다 */
function hwpxTextElementToString(tEl) {
  let s = '';
  tEl.childNodes.forEach((n) => {
    if (n.nodeType === 3 || n.nodeType === 4) s += n.nodeValue;
    else if (n.localName === 'lineBreak') s += '\n';
    else if (n.localName === 'tab' || n.localName === 'nbSpace' || n.localName === 'fwSpace') s += ' ';
  });
  return s;
}

/** 셀(<hp:tc>)의 글자. 문단은 줄바꿈으로 잇고, 셀 안에 든 중첩 표의 글자는 뺀다 */
function hwpxCellText(tcEl, tblEl) {
  const ownedBy = (n) => closestHwpxTable(n) === tblEl;
  return Array.from(tcEl.getElementsByTagNameNS('*', 'p'))
    .filter(ownedBy)
    .map((p) =>
      Array.from(p.getElementsByTagNameNS('*', 't'))
        .filter(ownedBy)
        .map(hwpxTextElementToString)
        .join('')
    )
    .join('\n')
    .trim();
}

/** <hp:tbl> 하나를 2차원 배열로 바꾼다. 각 셀은 cellAddr(행/열 주소)에 놓고, 병합 처리는 cellsToRows에 맡긴다 */
function hwpxTableToRows(tblEl) {
  const cells = [];
  childElementsByLocalName(tblEl, 'tr').forEach((trEl, trIdx) => {
    let colCursor = 0;
    childElementsByLocalName(trEl, 'tc').forEach((tcEl) => {
      const addr = childElementsByLocalName(tcEl, 'cellAddr')[0];
      const span = childElementsByLocalName(tcEl, 'cellSpan')[0];
      const row = intAttribute(addr, 'rowAddr', trIdx);
      const col = intAttribute(addr, 'colAddr', colCursor);
      const rowSpan = Math.max(1, intAttribute(span, 'rowSpan', 1));
      const colSpan = Math.max(1, intAttribute(span, 'colSpan', 1));
      colCursor = col + colSpan;
      cells.push({ row, col, rowSpan, colSpan, text: hwpxCellText(tcEl, tblEl) });
    });
  });
  return cellsToRows(cells);
}

/** .hwpx/.hwpt 문서의 모든 표를 문서에 나오는 순서대로 [{ rows }]로 읽는다 */
function readHwpxTables(filePath) {
  const zip = new PizZip(safeReadFileSync(filePath));
  const sectionNames = getHwpxSectionNames(zip);
  if (sectionNames.length === 0) {
    throw new Error('올바른 .hwpx 파일이 아닙니다. (Contents/section*.xml 없음)');
  }

  const tables = [];
  sectionNames.forEach((name) => {
    const doc = new DOMParser().parseFromString(zip.file(name).asText(), 'application/xml');
    if (doc.getElementsByTagName('parsererror').length > 0) {
      throw new Error(`한글 문서 본문(${name})을 해석할 수 없습니다. 파일이 손상되었을 수 있습니다.`);
    }
    Array.from(doc.getElementsByTagNameNS('*', 'tbl')).forEach((tblEl) => {
      tables.push({ rows: hwpxTableToRows(tblEl) });
    });
  });
  return tables;
}

/**
 * 한글 문서(.hwpx/.hwp)에서 읽은 표들을 "표 1", "표 2"... 이름의 시트로 담은 SheetJS
 * 워크북으로 만든다. 글자가 하나도 없는 빈 표는 뺀다.
 */
function buildWorkbookFromTables(rawTables) {
  const tables = rawTables
    .filter(({ rows }) => rows.some((r) => r.some((v) => v !== '')))
    .map(({ rows }, i) => ({ name: `표 ${i + 1}`, rows }));
  if (tables.length === 0) {
    throw new Error('한글 문서에서 표를 찾을 수 없습니다. 데이터가 표(행/열) 형태로 들어 있는지 확인하세요.');
  }

  const workbook = XLSX.utils.book_new();
  tables.forEach(({ name, rows }) => {
    // 빈 칸은 셀을 만들지 않아야(null) 완전히 빈 행이 엑셀처럼 "빈 행"으로 취급되어 데이터에서 빠진다.
    const sheet = XLSX.utils.aoa_to_sheet(rows.map((r) => r.map((v) => (v === '' ? null : v))));
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  });

  // 결재란·제목 상자 같은 작은 표가 데이터 표보다 먼저 나오는 경우가 많아, 행이 가장 많은 표를 기본으로 고른다.
  workbook.preferredSheetName = tables.reduce((best, t) => (t.rows.length > best.rows.length ? t : best)).name;
  workbook.sheetUnitLabel = '표';

  const describe = ({ name, rows }) => {
    const firstFilledRow = rows.find((r) => r.some((v) => v !== '')) || [];
    const preview = firstFilledRow.filter((v) => v !== '').slice(0, 3).join(', ').replace(/\s+/g, ' ');
    return `${name}(${rows.length}행×${rows[0].length}열: ${preview.length > 30 ? `${preview.slice(0, 30)}…` : preview})`;
  };
  log(`한글 문서에서 표 ${tables.length}개를 찾았습니다: ${tables.map(describe).join(' / ')}`);
  return workbook;
}

/**
 * 워크북 파싱 캐시. 같은 파일을 (시트 목록 조회 → 구조 분석 → 변환 실행) 단계마다
 * 매번 디스크에서 다시 읽고 다시 파싱하는 중복을 없애기 위함. 파일의 mtime이
 * 바뀌면(사용자가 외부에서 파일을 다시 저장한 경우) 자동으로 무효화된다.
 * 세션 중 여러 파일을 오가며 시험해볼 수 있으므로 무한정 커지지 않도록 개수를 제한한다.
 */
const workbookCache = new Map(); // filePath -> { mtimeMs, workbook }
const WORKBOOK_CACHE_LIMIT = 8;

/**
 * 파일 경로로부터 워크북을 읽는다.
 * - .hwpx/.hwpt/.hwp: 한글 문서 안의 표 하나하나를 시트 하나로 변환 (.hwp는 hwp-reader.js)
 * - 그 밖에는 확장자가 아니라 파일 앞부분(내용)으로 판단: 글자 파일(CSV)은 인코딩을 자동 판별
 *   (또는 encoding으로 직접 지정)해 문자열로, 그 외(xlsx·xls·HTML 표)는 SheetJS로 읽는다.
 * 결과 워크북의 fileInfo에 내용 형식·인코딩·확장자와 다른지·1904 날짜 체계 여부를 담는다.
 */
function readWorkbookSmart(filePath, { encoding = 'auto' } = {}) {
  const mtimeMs = getFileMtimeMs(filePath);
  const cacheKey = `${filePath}|${encoding}`;
  const cached = workbookCache.get(cacheKey);
  if (cached && cached.mtimeMs === mtimeMs) {
    return cached.workbook;
  }

  const ext = path.extname(filePath).toLowerCase();
  let workbook;
  // fileInfo: 인식 결과 화면에 보일 파일 정보(실제 내용 형식, 글자 인코딩, 확장자와 다른지)
  const fileInfo = { ext, contentKind: null, mismatch: '', encoding: null, encodingAuto: true, brokenChars: 0, textSource: false };
  if (ext === '.hwp') {
    workbook = buildWorkbookFromTables(readHwpTables(safeReadFileSync(filePath)));
    fileInfo.textSource = true;
  } else if (ext === '.hwpx' || ext === '.hwpt') {
    workbook = buildWorkbookFromTables(readHwpxTables(filePath));
    fileInfo.textSource = true;
  } else {
    const buffer = safeReadFileSync(filePath);
    // 확장자가 아니라 실제 내용으로 읽는 방법을 정한다(CSV를 .xlsx로 이름만 바꾼 파일 등)
    fileInfo.contentKind = tableReader.sniffContentKind(buffer);
    fileInfo.mismatch = tableReader.describeContentMismatch(ext, fileInfo.contentKind);
    if (fileInfo.contentKind === 'text') {
      const decoded = tableReader.decodeTextBuffer(buffer, encoding);
      Object.assign(fileInfo, { encoding: decoded.encoding, encodingAuto: decoded.auto, brokenChars: decoded.brokenChars, textSource: true });
      // raw:true — CSV 글자를 숫자/날짜로 추측 변환하지 않는다. 변환하면 전화번호
      // '01012345678'·우편번호 '03187'의 앞자리 0이 사라지고 '2026-01-15'가 날짜
      // 일련번호(46037)로 바뀐다. 날짜·숫자 글자는 탭1에서 열 단위로 확인해 바꾼다.
      workbook = XLSX.read(decoded.text, { type: 'string', raw: true, dense: true });
    } else {
      // 날짜는 cellDates 대신 readCellValue에서 직접 푼다(이유는 그 함수 설명 참고). cellNF는 셀의
      // 표시 형식(z)을 남겨 날짜 셀·앞자리 0 채움 셀을 알아보기 위해 필요하다.
      workbook = XLSX.read(buffer, { type: 'buffer', cellNF: true, dense: true });
    }
  }
  fileInfo.date1904 = tableReader.isDate1904(workbook);
  workbook.fileInfo = fileInfo;

  evictOldestIfFull(workbookCache, WORKBOOK_CACHE_LIMIT);
  workbookCache.set(cacheKey, { mtimeMs, workbook });
  return workbook;
}

document.getElementById('btnClearLog').addEventListener('click', () => {
  logArea.value = '';
});

// 표 형태 데이터 파일: xlsx와 매크로 포함 버전(xlsm)은 zip+XML 구조가 동일해 xlsx 파서로 함께 읽히고,
// 한글 문서(hwpx/hwp)와 그 서식 파일(hwpt)은 문서 안의 표를 시트처럼 읽는다
const DATA_FILE_FILTERS = [
  { name: 'Excel / CSV / 한글 문서 속 표', extensions: ['xlsx', 'xls', 'xlsm', 'csv', 'hwpx', 'hwpt', 'hwp'] },
  { name: 'All Files', extensions: ['*'] },
];

const DOCX_FILTERS = [
  { name: 'Word Document', extensions: ['docx'] },
  { name: 'All Files', extensions: ['*'] },
];

const HWPX_FILTERS = [
  { name: '한글(HWPX) 문서', extensions: ['hwpx'] },
  { name: 'All Files', extensions: ['*'] },
];

// 문서 계열: 각 본문 포맷(docx/hwpx)과 그 서식파일 버전(dotx/hwpt)은 내부 구조가 동일해 같은 엔진으로 처리됨
const TEMPLATE_FILTERS = [
  { name: 'Word / 한글(HWPX) 문서', extensions: ['docx', 'dotx', 'hwpx', 'hwpt'] },
  { name: 'All Files', extensions: ['*'] },
];


/**
 * 드롭존(클릭 + 드래그앤드롭) 공통 바인딩
 * onFileSelected(filePath) 콜백으로 선택된 파일 경로를 전달
 */
function bindDropzone(zoneEl, fileNameEl, filters, onFileSelected) {
  zoneEl.addEventListener('click', async () => {
    try {
      const paths = await ipcRenderer.invoke('dialog:openFile', { filters, multi: false });
      if (paths && paths.length > 0) {
        fileNameEl.textContent = path.basename(paths[0]);
        onFileSelected(paths[0]);
      }
    } catch (err) {
      notifyError('파일 선택 중 오류가 발생했습니다', err);
    }
  });

  zoneEl.addEventListener('dragover', (e) => {
    e.preventDefault();
    zoneEl.classList.add('dragover');
  });

  zoneEl.addEventListener('dragleave', () => {
    zoneEl.classList.remove('dragover');
  });

  zoneEl.addEventListener('drop', (e) => {
    e.preventDefault();
    zoneEl.classList.remove('dragover');
    try {
      const files = e.dataTransfer.files;
      if (!files || files.length === 0) return;
      // Electron 32부터 File.path가 없어져 webUtils로 실제 경로를 얻는다
      const filePath = webUtils.getPathForFile(files[0]);
      if (!filePath) {
        throw new Error('드래그된 파일의 경로를 확인할 수 없습니다.');
      }
      fileNameEl.textContent = path.basename(filePath);
      onFileSelected(filePath);
    } catch (err) {
      notifyError('파일 드롭 처리 중 오류가 발생했습니다', err);
    }
  });
}

/**
 * 파일 선택 직후의 최초 시트 반영은 onChange로, 사용자가 드롭다운에서 직접
 * 시트를 바꿀 때의 후속 처리(상태 초기화 등)는 onUserChangedSheet으로 분리
 * 되어 있다. '시트를 바꾸면 무엇을 초기화할지'는 호출하는 쪽의 책임이다.
 */
function refreshSheetSelector(filePath, sheetRowEl, sheetSelectEl, onChange, onUserChangedSheet) {
  try {
    const workbook = readWorkbookSmart(filePath);
    const sheetNames = workbook.SheetNames || [];
    sheetSelectEl.innerHTML = '';
    sheetNames.forEach((name) => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sheetSelectEl.appendChild(opt);
    });

    // 한글 문서는 시트 대신 '표'를 고르므로 안내 문구도 그에 맞춘다
    const unit = workbook.sheetUnitLabel || '시트';
    const labelEl = sheetRowEl.querySelector('label');
    if (labelEl) labelEl.textContent = `${unit} 선택`;

    if (sheetNames.length > 1) {
      sheetRowEl.hidden = false;
      log(`${unit} ${sheetNames.length}개 발견: ${sheetNames.join(', ')}`);
    } else {
      sheetRowEl.hidden = true;
    }

    const selected = sheetNames.includes(workbook.preferredSheetName) ? workbook.preferredSheetName : sheetNames[0] || null;
    sheetSelectEl.value = selected || '';
    onChange(selected);

    sheetSelectEl.onchange = () => {
      onChange(sheetSelectEl.value);
      if (onUserChangedSheet) onUserChangedSheet(sheetSelectEl.value);
    };
  } catch (err) {
    sheetRowEl.hidden = true;
    notifyError('시트 목록을 읽는 중 오류가 발생했습니다', err);
  }
}

/**
 * 시트에서 헤더 행을 찾고(또는 직접 고른 행을 쓰고) 그 아래 행들을 { 컬럼키: 값 } 객체로 읽는다.
 * 자세한 규칙은 table-reader.js의 buildTableFromSheet. 반환값의 report에 인식 결과(헤더 행,
 * 건너뛴 행, 이름 없는 열, 인코딩, 확장자와 다른 내용)가 들어 있다.
 * options: { displayText, headerRowIndex(0부터 · null이면 자동), encoding('auto'|'utf-8'|'cp949') }
 */
function readSheetTable(filePath, sheetName, { displayText = false, headerRowIndex = null, encoding = 'auto' } = {}) {
  const workbook = readWorkbookSmart(filePath, { encoding });
  const targetSheetName = sheetName || workbook.SheetNames[0];
  if (!targetSheetName || !workbook.Sheets[targetSheetName]) {
    throw new Error(`시트를 찾을 수 없습니다${sheetName ? ` (${sheetName})` : ''}.`);
  }
  const fileInfo = workbook.fileInfo || {};
  const table = tableReader.buildTableFromSheet(workbook.Sheets[targetSheetName], targetSheetName, {
    displayText,
    date1904: !!fileInfo.date1904,
    headerRowIndex,
  });
  table.fileInfo = fileInfo;
  table.textSource = !!fileInfo.textSource;
  return table;
}

/* -------------------------------------------------------------------------
 * 컬럼 이름 자동 연결은 auto-match.js(순수 함수). 여기서는 사용자 동의어 사전을 넣어 부른다.
 * 템플릿 변수를 한꺼번에 정하므로 순서에 따라 결과가 달라지지 않도록 autoMatchAll을 쓴다.
 * ------------------------------------------------------------------------- */

const autoMatchAll = (headersB, headersA) => autoMatch.autoMatchAll(headersB, headersA, synonymGroups);

/** XML 텍스트 노드에 안전하게 삽입할 수 있도록 특수문자를 이스케이프 */
function xmlEscape(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/* -------------------------------------------------------------------------
 * 동의어 사전 (자동 매핑 2단계에서 사용, 사용자가 직접 추가/수정 가능)
 * "성명"↔"이름"처럼 표기는 다르지만 같은 의미인 컬럼명을, AI 모델 없이도
 * 가장 흔한 사례만큼은 바로 매핑하기 위한 가벼운 대안. 프리셋과 마찬가지로
 * 사용자 데이터 폴더에 저장되어 앱을 다시 실행해도 유지되며, 모달에서
 * "한 줄 = 동의어 한 그룹(쉼표로 구분)" 형태로 직접 편집할 수 있다.
 * 템플릿 변수 ↔ 데이터 컬럼 자동 매핑(autoMatchAll)이 쓴다.
 * ------------------------------------------------------------------------- */

const SYNONYM_GROUPS_FILE_NAME = 'synonym-groups.json';

const DEFAULT_SYNONYM_GROUPS = [
  ['이름', '성명', '성함'],
  ['생년월일', '출생일자', '생일'],
  ['전화번호', '연락처', '핸드폰번호', '휴대폰번호', '휴대전화'],
  ['이메일', '이메일주소', '메일주소', '전자우편'],
  ['주소', '소재지', '거주지', '주소지'],
  ['부서', '소속', '소속부서', '부서명'],
  ['직급', '직책', '직위'],
  ['나이', '연령'],
  ['회사명', '업체명', '상호', '회사'],
  ['담당자', '담당자명', '책임자'],
  ['성별', 'sex', 'gender'],
  ['비고', '메모', '참고사항'],
];

let synonymGroups = [];

function getSynonymGroupsFilePath() {
  return path.join(getUserDataPath(), SYNONYM_GROUPS_FILE_NAME);
}

function loadSynonymGroupsFromDisk() {
  try {
    if (!fs.existsSync(getSynonymGroupsFilePath())) {
      synonymGroups = DEFAULT_SYNONYM_GROUPS.map((g) => [...g]);
      saveSynonymGroupsToDisk();
      return;
    }
    const raw = safeReadFileSync(getSynonymGroupsFilePath(), 'utf-8');
    const parsed = JSON.parse(raw);
    synonymGroups = Array.isArray(parsed)
      ? parsed.filter((g) => Array.isArray(g) && g.length >= 2)
      : DEFAULT_SYNONYM_GROUPS.map((g) => [...g]);
  } catch (err) {
    synonymGroups = DEFAULT_SYNONYM_GROUPS.map((g) => [...g]);
    log(`동의어 사전을 불러오는 중 오류가 발생하여 기본값으로 되돌립니다: ${friendlyErrorMessage(err)}`, 'warn');
  }
}

function saveSynonymGroupsToDisk() {
  safeWriteFileSync(getSynonymGroupsFilePath(), JSON.stringify(synonymGroups, null, 2));
}

/** 동의어 그룹 배열 ↔ "한 줄 = 한 그룹(쉼표 구분)" 텍스트 상호 변환 */
function synonymGroupsToText(groups) {
  return groups.map((g) => g.join(', ')).join('\n');
}
function textToSynonymGroups(text) {
  return text
    .split('\n')
    .map((line) =>
      line
        .split(',')
        .map((term) => term.trim())
        .filter((term) => term !== '')
    )
    .filter((group) => group.length >= 2);
}

// 동의어 사전 관리 모달
const synonymManagerModal = document.getElementById('synonymManagerModal');
const synonymGroupsTextarea = document.getElementById('synonymGroupsTextarea');
const synonymManagerError = document.getElementById('synonymManagerError');
const synonymManagerSaveBtn = document.getElementById('synonymManagerSave');
const synonymManagerCancelBtn = document.getElementById('synonymManagerCancel');
const synonymManagerResetBtn = document.getElementById('synonymManagerReset');

function openSynonymManagerModal() {
  synonymGroupsTextarea.value = synonymGroupsToText(synonymGroups);
  synonymManagerError.hidden = true;
  synonymManagerModal.hidden = false;
  synonymGroupsTextarea.focus();
}

function closeSynonymManagerModal() {
  synonymManagerModal.hidden = true;
}

synonymManagerCancelBtn.addEventListener('click', closeSynonymManagerModal);
synonymManagerModal.addEventListener('click', (e) => {
  if (e.target === synonymManagerModal) closeSynonymManagerModal();
});
synonymManagerResetBtn.addEventListener('click', () => {
  if (!confirm('기본 동의어 사전으로 되돌릴까요? 직접 추가/수정한 내용은 사라집니다.')) return;
  synonymGroupsTextarea.value = synonymGroupsToText(DEFAULT_SYNONYM_GROUPS);
});
synonymManagerSaveBtn.addEventListener('click', () => {
  try {
    synonymGroups = textToSynonymGroups(synonymGroupsTextarea.value);
    saveSynonymGroupsToDisk();
    closeSynonymManagerModal();
    log(`동의어 사전을 저장했습니다 (${synonymGroups.length}개 그룹). 다음 [데이터 불러오기]부터 적용됩니다.`, 'ok');
  } catch (err) {
    synonymManagerError.textContent = friendlyErrorMessage(err);
    synonymManagerError.hidden = false;
  }
});

document.getElementById('btnManageSynonymsTab2').addEventListener('click', openSynonymManagerModal);

/* -------------------------------------------------------------------------
 * 사용자 데이터 폴더 경로(동의어 사전 저장 위치). 앱 시작 때 메인 프로세스에서 받아 온다.
 * ------------------------------------------------------------------------- */

let userDataPath = null;

function getUserDataPath() {
  return userDataPath;
}

// 앱 시작 시 사용자 데이터 폴더 경로를 받아와 동의어 사전을 불러온다
ipcRenderer
  .invoke('app:getUserDataPath')
  .then((p) => {
    userDataPath = p;
    loadSynonymGroupsFromDisk();
  })
  .catch((err) => {
    log(`앱 데이터 폴더 경로를 확인하지 못했습니다: ${friendlyErrorMessage(err)}`, 'warn');
  });

module.exports = {
  log,
  notifyError,
  friendlyErrorMessage,
  safeReadFileSync,
  safeWriteFileSync,
  getFileMtimeMs,
  evictOldestIfFull,
  confirmModal,
  readWorkbookSmart,
  getHwpxSectionNames,
  DATA_FILE_FILTERS,
  DOCX_FILTERS,
  HWPX_FILTERS,
  TEMPLATE_FILTERS,
  bindDropzone,
  refreshSheetSelector,
  readSheetTable,
  describeRecognition: tableReader.describeRecognition,
  autoMatchAll,
  xmlEscape,
  formatPreviewValue,
};
