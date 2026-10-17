// common.js
// 탭1(문서 양식 변환기)/탭2(메일머지)가 공통으로 쓰는 유틸·캐시·동의어 사전·
// 헤더 자동 감지·컬럼 자동 매핑 로직. 두 탭 모두 이 모듈을 require 하여 쓴다.

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const iconv = require('iconv-lite');
const { readHwpTables, cellsToRows } = require('./hwp-reader.js');

/* =========================================================================
   0. 공통 유틸 (로그, 탭 전환)
   ========================================================================= */

const logArea = document.getElementById('logArea');

function log(message, level = 'info') {
  const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
  const prefix = { info: 'INFO', ok: ' OK ', warn: 'WARN', error: 'FAIL' }[level] || 'INFO';
  logArea.value += `[${time}] [${prefix}] ${message}\n`;
  logArea.scrollTop = logArea.scrollHeight;
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

/**
 * 버퍼가 유효한 UTF-8 시퀀스로만 구성되어 있는지 검사한다.
 * (한글 CSV가 EUC-KR/CP949(Windows 기본 인코딩)로 저장된 경우를 구분하기 위함)
 */
function isValidUTF8(buffer) {
  let i = 0;
  const len = buffer.length;
  while (i < len) {
    const byte = buffer[i];
    let extraBytes = 0;

    if (byte <= 0x7f) {
      extraBytes = 0;
    } else if ((byte & 0xe0) === 0xc0) {
      extraBytes = 1;
    } else if ((byte & 0xf0) === 0xe0) {
      extraBytes = 2;
    } else if ((byte & 0xf8) === 0xf0) {
      extraBytes = 3;
    } else {
      return false;
    }

    if (i + extraBytes >= len) return false;

    for (let j = 1; j <= extraBytes; j++) {
      if ((buffer[i + j] & 0xc0) !== 0x80) return false;
    }

    i += extraBytes + 1;
  }
  return true;
}

/**
 * CSV 파일을 인코딩을 자동 판별하여 문자열로 디코딩한다.
 * - UTF-8 BOM: BOM 제거 후 UTF-8로 디코딩
 * - 유효한 UTF-8: 그대로 UTF-8로 디코딩
 * - 그 외(대부분 Windows 한글 Excel의 기본 저장 인코딩): CP949(EUC-KR 상위 호환)로 디코딩
 */
function decodeCsvBuffer(buffer) {
  const hasBOM = buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
  if (hasBOM) {
    return buffer.slice(3).toString('utf-8');
  }
  if (isValidUTF8(buffer)) {
    return buffer.toString('utf-8');
  }
  return iconv.decode(buffer, 'cp949');
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
 * 반환한다. 탭1의 변환 결과 확인 모달과 탭2의 일괄 생성 확인 모달이 이 구조를
 * 그대로 공유하므로 여기만 한 번 구현한다 — 모달에 채울 내용(요약/미리보기
 * 테이블)을 만드는 로직은 서로 달라 호출하는 쪽에 남겨둔다.
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
 * - .csv: 인코딩 자동 판별 후 문자열로 파싱 (한글 CP949 CSV 대응)
 * - .hwpx/.hwpt/.hwp: 한글 문서 안의 표 하나하나를 시트 하나로 변환 (.hwp는 hwp-reader.js)
 * - .xlsx/.xls: XLSX.readFile 사용
 */
function readWorkbookSmart(filePath) {
  const mtimeMs = getFileMtimeMs(filePath);
  const cached = workbookCache.get(filePath);
  if (cached && cached.mtimeMs === mtimeMs) {
    return cached.workbook;
  }

  const ext = path.extname(filePath).toLowerCase();
  let workbook;
  if (ext === '.hwp') {
    workbook = buildWorkbookFromTables(readHwpTables(safeReadFileSync(filePath)));
  } else if (ext === '.hwpx' || ext === '.hwpt') {
    workbook = buildWorkbookFromTables(readHwpxTables(filePath));
  } else if (ext === '.csv') {
    const buffer = safeReadFileSync(filePath);
    const text = decodeCsvBuffer(buffer);
    workbook = XLSX.read(text, { type: 'string', raw: false });
  } else {
    const buffer = safeReadFileSync(filePath);
    workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  }

  evictOldestIfFull(workbookCache, WORKBOOK_CACHE_LIMIT);
  workbookCache.set(filePath, { mtimeMs, workbook });
  return workbook;
}

document.getElementById('btnClearLog').addEventListener('click', () => {
  logArea.value = '';
});

// 탭 전환
document.querySelectorAll('.tab-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach((p) => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.tab).classList.add('active');
  });
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

const JSON_FILTERS = [
  { name: 'JSON Files', extensions: ['json'] },
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
      const filePath = files[0].path;
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
 * 되어 있다. 탭1(A/B 파일)과 탭2(메일머지 데이터 파일)가 이 함수를 공유하므로,
 * "시트를 바꾸면 무엇을 초기화할지"는 호출하는 쪽의 책임으로 둬야 한다 —
 * 그렇지 않으면 탭2에서 시트를 바꿨는데 탭1의 매핑/프리셋이 초기화되는 식의
 * 엉뚱한 부작용이 생긴다.
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

    // 한글 문서는 시트 대신 "표"를 고르므로 안내 문구도 그에 맞춘다
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
 * 행(배열) 하나가 "헤더(컬럼명) 행"일 가능성이 얼마나 되는지 점수로 매긴다.
 * - 완전히 빈 행: 후보 아님(-1)
 * - 여러 칸짜리 표인데 딱 한 칸만 채워진 행: "2026년 명단"류 제목 행일 가능성이
 *   높으므로 낮은 점수(0)
 * - 그 외에는 채워진 칸 수가 많을수록 높은 점수. 평균 글자 수가 아주 길면
 *   (문장형 텍스트) 헤더보다는 본문/설명일 가능성이 있어 약하게 감점.
 */
function scoreHeaderRowCandidate(row) {
  const nonEmpty = row.filter((c) => String(c).trim() !== '');
  if (nonEmpty.length === 0) return -1;
  if (nonEmpty.length === 1 && row.length > 1) return 0;
  const avgLen = nonEmpty.reduce((sum, c) => sum + String(c).trim().length, 0) / nonEmpty.length;
  return avgLen > 20 ? nonEmpty.length * 0.5 : nonEmpty.length;
}

/**
 * 시트 상단 최대 10행을 살펴보고 실제 헤더 행으로 보이는 행의 인덱스(0-based,
 * 시트의 실제 행 번호와 일치)를 추정한다. "2026년 하반기 명단"처럼 제목 행이
 * 1행에 있어 헤더가 그 아래에 있는 경우를 걸러내기 위함. 점수가 동일하면 더
 * 앞쪽 행(기존 동작인 1행)을 우선하므로, 평범한 파일에서는 항상 1행이 그대로
 * 선택된다.
 */
function detectHeaderRowIndex(rawRows) {
  const scanLimit = Math.min(10, rawRows.length);
  let bestIndex = 0;
  let bestScore = -Infinity;
  for (let i = 0; i < scanLimit; i++) {
    const score = scoreHeaderRowCandidate(rawRows[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestIndex;
}

/**
 * 엑셀/CSV 파일의 지정된 시트에서 헤더 행을 자동으로 찾아 배열로 추출한다.
 * { headers, headerRowIndex } 를 반환하며, headerRowIndex는 변환 시 데이터
 * 행을 정확히 그 아래부터 읽기 위해 호출한 쪽에서 보관해 두어야 한다.
 */
function extractHeaders(filePath, sheetName) {
  const workbook = readWorkbookSmart(filePath);
  const targetSheetName = sheetName || workbook.SheetNames[0];
  if (!targetSheetName || !workbook.Sheets[targetSheetName]) {
    throw new Error(`시트를 찾을 수 없습니다${sheetName ? ` (${sheetName})` : ''}.`);
  }
  const sheet = workbook.Sheets[targetSheetName];
  // blankrows:false를 쓰지 않는다 — 빈 행을 걸러내면 배열 인덱스가 실제 시트
  // 행 번호와 어긋나서, 아래에서 구한 headerRowIndex를 이후 데이터 읽기(range)에
  // 그대로 재사용할 수 없게 된다.
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rawRows || rawRows.length === 0) throw new Error(`"${targetSheetName}" 시트에 데이터가 비어 있습니다.`);

  const headerRowIndex = detectHeaderRowIndex(rawRows);
  const headerRow = rawRows[headerRowIndex].map((h) => String(h).trim()).filter((h) => h !== '');
  if (headerRow.length === 0) throw new Error(`"${targetSheetName}" 시트에서 헤더(컬럼명) 행을 찾을 수 없습니다.`);
  return { headers: headerRow, headerRowIndex };
}

/** 두 문자열 간 편집거리(Levenshtein distance)를 계산 */
function levenshteinDistance(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
    }
    prev = cur;
  }
  return prev[n];
}

const FUZZY_MAX_DISTANCE = 2; // 절대 편집거리 상한
const FUZZY_MAX_RATIO = 0.34; // 상대 편집거리 상한(짧은 이름에서 과매칭 방지)

/** 문자열 비교용 정규화(공백 제거 + 소문자 변환). 자동 매핑/동의어 사전에서 공통으로 사용 */
function normalizeFieldName(s) {
  return String(s).replace(/\s+/g, '').toLowerCase();
}

/* -------------------------------------------------------------------------
 * 한글 자모 분해 (자동 매핑 3-1순위: 음절 단위 퍼지 매칭이 놓치는 한글 오타 보완)
 * "성명"↔"승명"처럼 음절 하나가 통째로 달라 보이지만 실제로는 자모 하나(모음
 * ㅓ↔ㅡ 등)만 다른 흔한 오타가 있다. 음절 단위 편집거리는 이런 경우 "음절 1개가
 * 통째로 바뀜"으로 보기 때문에, 짧은 한글 필드명(2~4음절)에서는 상대 편집거리
 * 비율이 쉽게 FUZZY_MAX_RATIO를 넘어 매칭되지 못한다. 음절을 초성/중성/종성
 * 자모로 풀어서 같은 거리 계산을 다시 하면, 음절 1개 중 자모 1개만 다른 경우의
 * 비율이 훨씬 작게 나와 "매우 비슷함"으로 올바르게 인식된다. 한글이 아닌 문자는
 * 그대로 통과되므로 영문/숫자 필드명의 매칭 결과에는 영향이 없다.
 * ------------------------------------------------------------------------- */

const HANGUL_CHOSEONG = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'];
const HANGUL_JUNGSEONG = ['ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ', 'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ'];
const HANGUL_JONGSEONG = ['', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ', 'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'];
const HANGUL_SYLLABLE_BASE = 0xac00;
const HANGUL_SYLLABLE_LAST = 0xd7a3;

/** 한글 완성형 음절을 초성/중성/종성 자모로 분해한다. 한글이 아닌 문자는 그대로 둔다. */
function decomposeHangulToJamo(str) {
  let result = '';
  for (const ch of str) {
    const code = ch.codePointAt(0);
    if (code >= HANGUL_SYLLABLE_BASE && code <= HANGUL_SYLLABLE_LAST) {
      const offset = code - HANGUL_SYLLABLE_BASE;
      const cho = Math.floor(offset / (21 * 28));
      const jung = Math.floor((offset % (21 * 28)) / 28);
      const jong = offset % 28;
      result += HANGUL_CHOSEONG[cho] + HANGUL_JUNGSEONG[jung] + HANGUL_JONGSEONG[jong];
    } else {
      result += ch;
    }
  }
  return result;
}

const JAMO_FUZZY_MAX_DISTANCE = 2; // 절대 편집거리 상한(자모 단위)
const JAMO_FUZZY_MAX_RATIO = 0.3; // 음절 단위(0.34)보다 더 엄격하게 잡아 과매칭을 방지

/**
 * headersA 중 bCol과 가장 비슷한 컬럼을 찾는다.
 * 1순위: 정규화 후 완전히 같은 이름("exact").
 * 2순위: 사용자가 관리하는 동의어 사전에 bCol과 같은 그룹으로 등록된 이름이
 *   headersA에 유일하게 하나 있을 때만 "synonym"으로 매칭한다(예: "성명"↔"이름").
 *   후보가 여러 개면 모호하므로 매칭하지 않고, 0개면 사전에 없는 것일 뿐이니
 *   다음 단계(퍼지)를 계속 시도한다.
 * 3순위: 편집거리가 절대/상대 기준 모두를 만족하면서 유일하게 가장 가까운
 *   후보가 있을 때만 "fuzzy"로 매칭한다.
 * 4순위: 3순위에서 못 찾았을 때만, 한글 음절을 자모로 풀어 같은 방식의 편집거리를
 *   다시 계산한다("성명"↔"승명"처럼 음절 1개 중 자모 1개만 다른 한글 오타를
 *   잡아낸다). 한글이 아닌 이름은 자모 분해가 원문 그대로 통과되므로 3순위와
 *   동일한 결과가 나와 새로 매칭되는 일이 없다.
 * 모든 단계에서 후보가 여러 개 동률이면 모호하므로 매칭하지 않는다 — 잘못
 * 자동 매핑되어 데이터가 다른 컬럼에 들어가는 것이 매핑이 안 되는 것보다
 * 더 나쁘다. usedHeadersA에 담긴 컬럼은 이미 다른 B 컬럼에 매칭되었으므로
 * 후보에서 제외해 하나의 A 컬럼이 여러 B 컬럼에 중복 자동매핑되지 않도록 한다.
 */
function findAutoMatch(bCol, headersA, usedHeadersA) {
  const normalize = normalizeFieldName;
  const bNorm = normalize(bCol);
  const available = headersA.filter((a) => !usedHeadersA.has(a));

  const exact = available.find((a) => normalize(a) === bNorm);
  if (exact) return { match: exact, reason: 'exact' };

  const group = synonymGroups.find((g) => g.some((term) => normalize(term) === bNorm));
  if (group) {
    const groupNormSet = new Set(group.map(normalize));
    const synonymCandidates = available.filter((a) => groupNormSet.has(normalize(a)));
    if (synonymCandidates.length === 1) {
      return { match: synonymCandidates[0], reason: 'synonym' };
    }
    if (synonymCandidates.length > 1) {
      return { match: null, reason: null }; // 동의어 사전상으로도 모호함
    }
    // 0건이면 이 사전 그룹에 해당하는 이름이 데이터에 없는 것뿐이므로 퍼지 매칭을 계속 시도
  }

  const candidates = available
    .map((a) => {
      const aNorm = normalize(a);
      const dist = levenshteinDistance(aNorm, bNorm);
      const ratio = dist / Math.max(aNorm.length, bNorm.length, 1);
      return { a, dist, ratio };
    })
    .filter((c) => c.dist > 0 && c.dist <= FUZZY_MAX_DISTANCE && c.ratio <= FUZZY_MAX_RATIO)
    .sort((x, y) => x.dist - y.dist);

  if (candidates.length > 0 && (candidates.length === 1 || candidates[0].dist < candidates[1].dist)) {
    return { match: candidates[0].a, reason: 'fuzzy' };
  }

  const bJamo = decomposeHangulToJamo(bNorm);
  const jamoCandidates = available
    .map((a) => {
      const aJamo = decomposeHangulToJamo(normalize(a));
      const dist = levenshteinDistance(aJamo, bJamo);
      const ratio = dist / Math.max(aJamo.length, bJamo.length, 1);
      return { a, dist, ratio };
    })
    .filter((c) => c.dist > 0 && c.dist <= JAMO_FUZZY_MAX_DISTANCE && c.ratio <= JAMO_FUZZY_MAX_RATIO)
    .sort((x, y) => x.dist - y.dist);

  if (jamoCandidates.length > 0 && (jamoCandidates.length === 1 || jamoCandidates[0].dist < jamoCandidates[1].dist)) {
    return { match: jamoCandidates[0].a, reason: 'fuzzy' };
  }

  return { match: null, reason: null };
}

/**
 * 값을 화면 표시용 문자열로 안전하게 바꾼다. 엑셀 날짜 셀은 readWorkbookSmart가
 * cellDates:true로 읽어 JS Date 객체가 되는데, 그냥 String()으로 바꾸면
 * "Thu Jan 01 2026 00:00:00 GMT+0900..." 같은 장문이 되어버리므로, 시/분/초가
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
 * 탭1(컬럼 매핑)과 탭2(메일머지 변수 매핑)가 findAutoMatch를 통해 함께 쓴다.
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
    log(`동의어 사전을 저장했습니다 (${synonymGroups.length}개 그룹). 다음 [구조 분석]/[데이터 불러오기]부터 적용됩니다.`, 'ok');
  } catch (err) {
    synonymManagerError.textContent = friendlyErrorMessage(err);
    synonymManagerError.hidden = false;
  }
});

[document.getElementById('btnManageSynonymsTab1'), document.getElementById('btnManageSynonymsTab2')].forEach((btn) => {
  if (btn) btn.addEventListener('click', openSynonymManagerModal);
});

/* -------------------------------------------------------------------------
 * 사용자 데이터 폴더 경로 (프리셋은 탭1, 동의어 사전은 이 모듈이 사용).
 * 값이 준비되기 전에 다른 모듈이 먼저 필요로 할 수 있으므로, 직접 변수를
 * 내보내는 대신 getUserDataPath()/onUserDataPathReady() 접근자로 노출한다
 * — require 시점에 값을 복사해버리는 구조화 할당(destructuring)은 이후
 * 비동기로 채워지는 값을 반영하지 못하기 때문이다.
 * ------------------------------------------------------------------------- */

let userDataPath = null;
const userDataPathReadyCallbacks = [];

function getUserDataPath() {
  return userDataPath;
}

function onUserDataPathReady(callback) {
  if (userDataPath !== null) callback(userDataPath);
  else userDataPathReadyCallbacks.push(callback);
}

// 앱 시작 시 사용자 데이터 폴더 경로를 받아와 동의어 사전을 불러오고,
// 이를 기다리고 있던 다른 모듈(탭1의 프리셋 로드 등)에도 알린다.
ipcRenderer
  .invoke('app:getUserDataPath')
  .then((p) => {
    userDataPath = p;
    loadSynonymGroupsFromDisk();
    userDataPathReadyCallbacks.forEach((cb) => cb(p));
    userDataPathReadyCallbacks.length = 0;
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
  JSON_FILTERS,
  bindDropzone,
  refreshSheetSelector,
  extractHeaders,
  normalizeFieldName,
  levenshteinDistance,
  decomposeHangulToJamo,
  findAutoMatch,
  xmlEscape,
  formatPreviewValue,
  getUserDataPath,
  onUserDataPathReady,
};
