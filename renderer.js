// renderer.js
// 렌더러 프로세스: 탭 UI 제어, 엑셀 구조분석/변환, 메일머지 변수추출/치환 로직

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const iconv = require('iconv-lite');

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

/**
 * 파일 경로로부터 워크북을 읽는다.
 * - .csv: 인코딩 자동 판별 후 문자열로 파싱 (한글 CP949 CSV 대응)
 * - .xlsx/.xls: XLSX.readFile 사용
 */
function readWorkbookSmart(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.csv') {
    const buffer = safeReadFileSync(filePath);
    const text = decodeCsvBuffer(buffer);
    return XLSX.read(text, { type: 'string', raw: false });
  }
  const buffer = safeReadFileSync(filePath);
  return XLSX.read(buffer, { type: 'buffer', cellDates: true });
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

const EXCEL_FILTERS = [
  { name: 'Excel / CSV Files', extensions: ['xlsx', 'xls', 'csv'] },
  { name: 'All Files', extensions: ['*'] },
];

const DOCX_FILTERS = [
  { name: 'Word Document', extensions: ['docx'] },
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

/* =========================================================================
   TAB 1. 문서 양식 변환기
   ========================================================================= */

let fileAPath = null;
let fileBPath = null;
let sheetAName = null;
let sheetBName = null;
let headersA = [];
let headersB = [];
let mapping = {}; // { bColumn: aColumn|null }

const dropA = document.getElementById('dropA');
const dropB = document.getElementById('dropB');
const fileNameA = document.getElementById('fileNameA');
const fileNameB = document.getElementById('fileNameB');
const sheetRowA = document.getElementById('sheetRowA');
const sheetRowB = document.getElementById('sheetRowB');
const sheetSelectA = document.getElementById('sheetSelectA');
const sheetSelectB = document.getElementById('sheetSelectB');
const analyzeStatus = document.getElementById('analyzeStatus');
const mappingTableBody = document.getElementById('mappingTableBody');

// 시트 선택 <select> 클릭이 드롭존의 파일선택 클릭으로 전파되지 않도록 차단
[sheetRowA, sheetRowB].forEach((row) => {
  row.addEventListener('click', (e) => e.stopPropagation());
});

/** 파일이 선택되면 시트 목록을 읽어 시트 선택 드롭다운을 갱신한다 */
function refreshSheetSelector(filePath, sheetRowEl, sheetSelectEl, onChange) {
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

    if (sheetNames.length > 1) {
      sheetRowEl.hidden = false;
      log(`시트 ${sheetNames.length}개 발견: ${sheetNames.join(', ')}`);
    } else {
      sheetRowEl.hidden = true;
    }

    const selected = sheetNames[0] || null;
    onChange(selected);

    sheetSelectEl.onchange = () => {
      onChange(sheetSelectEl.value);
      resetAnalysisState();
      log(`시트가 "${sheetSelectEl.value}"(으)로 변경되었습니다. [구조 분석]을 다시 실행하세요.`, 'warn');
    };
  } catch (err) {
    sheetRowEl.hidden = true;
    notifyError('시트 목록을 읽는 중 오류가 발생했습니다', err);
  }
}

/** 파일이 (재)선택되면 이전 분석 결과가 새 파일과 어긋나지 않도록 매핑 상태를 초기화 */
function resetAnalysisState() {
  headersA = [];
  headersB = [];
  mapping = {};
  renderMappingTable();
  analyzeStatus.textContent = '대기 중';
  analyzeStatus.className = 'pill warn';
}

bindDropzone(dropA, fileNameA, EXCEL_FILTERS, (p) => {
  fileAPath = p;
  log(`원본(A) 파일 선택됨: ${p}`);
  resetAnalysisState();
  refreshSheetSelector(p, sheetRowA, sheetSelectA, (name) => {
    sheetAName = name;
  });
});

bindDropzone(dropB, fileNameB, EXCEL_FILTERS, (p) => {
  fileBPath = p;
  log(`타겟(B) 파일 선택됨: ${p}`);
  resetAnalysisState();
  refreshSheetSelector(p, sheetRowB, sheetSelectB, (name) => {
    sheetBName = name;
  });
});

/** 엑셀/CSV 파일의 지정된 시트에서 첫 번째 행(헤더)을 배열로 추출 */
function extractHeaders(filePath, sheetName) {
  const workbook = readWorkbookSmart(filePath);
  const targetSheetName = sheetName || workbook.SheetNames[0];
  if (!targetSheetName || !workbook.Sheets[targetSheetName]) {
    throw new Error(`시트를 찾을 수 없습니다${sheetName ? ` (${sheetName})` : ''}.`);
  }
  const sheet = workbook.Sheets[targetSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  if (!rows || rows.length === 0) throw new Error(`"${targetSheetName}" 시트에 데이터가 비어 있습니다.`);
  const headerRow = rows[0].map((h) => String(h).trim()).filter((h) => h !== '');
  if (headerRow.length === 0) throw new Error(`"${targetSheetName}" 시트에서 헤더(컬럼명) 행을 찾을 수 없습니다.`);
  return headerRow;
}

function renderMappingTable() {
  mappingTableBody.innerHTML = '';
  mapping = {};

  if (headersB.length === 0) {
    mappingTableBody.innerHTML = '<tr><td colspan="2" class="empty-hint">먼저 [구조 분석]을 실행하세요.</td></tr>';
    return;
  }

  headersB.forEach((bCol) => {
    const tr = document.createElement('tr');

    const tdB = document.createElement('td');
    tdB.textContent = bCol;
    tr.appendChild(tdB);

    const tdSelect = document.createElement('td');
    const select = document.createElement('select');
    select.dataset.bcolumn = bCol;

    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = '-- 선택 안함 --';
    select.appendChild(emptyOpt);

    headersA.forEach((aCol) => {
      const opt = document.createElement('option');
      opt.value = aCol;
      opt.textContent = aCol;
      select.appendChild(opt);
    });

    // 이름이 동일하거나 유사하면 자동 매핑 (편의 기능)
    const normalize = (s) => s.replace(/\s+/g, '').toLowerCase();
    const autoMatch = headersA.find((a) => normalize(a) === normalize(bCol));
    if (autoMatch) {
      select.value = autoMatch;
      mapping[bCol] = autoMatch;
    } else {
      mapping[bCol] = '';
    }

    select.addEventListener('change', () => {
      mapping[bCol] = select.value;
    });

    tdSelect.appendChild(select);
    tr.appendChild(tdSelect);
    mappingTableBody.appendChild(tr);
  });
}

document.getElementById('btnAnalyze').addEventListener('click', () => {
  try {
    if (!fileAPath || !fileBPath) {
      throw new Error('원본(A) 파일과 타겟 양식(B) 파일을 모두 선택하세요.');
    }
    log(`구조 분석을 시작합니다... (A 시트: ${sheetAName || '첫 번째 시트'}, B 시트: ${sheetBName || '첫 번째 시트'})`);
    headersA = extractHeaders(fileAPath, sheetAName);
    headersB = extractHeaders(fileBPath, sheetBName);
    log(`원본(A) 컬럼 ${headersA.length}개, 타겟(B) 컬럼 ${headersB.length}개 발견`, 'ok');

    renderMappingTable();
    analyzeStatus.textContent = '분석 완료';
    analyzeStatus.className = 'pill ok';
    log('컬럼 매핑 테이블이 생성되었습니다.', 'ok');
  } catch (err) {
    analyzeStatus.textContent = '분석 실패';
    analyzeStatus.className = 'pill warn';
    notifyError('구조 분석 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnSaveMapping').addEventListener('click', async () => {
  try {
    if (headersB.length === 0) throw new Error('저장할 매핑 규칙이 없습니다. 먼저 구조 분석을 실행하세요.');

    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '매핑 규칙 저장',
      defaultPath: 'mapping_rule.json',
      filters: JSON_FILTERS,
    });
    if (!savePath) return;

    const payload = {
      sourceHeaders: headersA,
      targetHeaders: headersB,
      sourceSheet: sheetAName,
      targetSheet: sheetBName,
      mapping,
    };
    safeWriteFileSync(savePath, JSON.stringify(payload, null, 2));
    log(`매핑 규칙이 저장되었습니다: ${savePath}`, 'ok');
  } catch (err) {
    notifyError('매핑 규칙 저장 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnLoadMapping').addEventListener('click', async () => {
  try {
    const paths = await ipcRenderer.invoke('dialog:openFile', { filters: JSON_FILTERS, multi: false });
    if (!paths || paths.length === 0) return;

    const raw = safeReadFileSync(paths[0], 'utf-8');
    let payload;
    try {
      payload = JSON.parse(raw);
    } catch (parseErr) {
      throw new Error('올바른 JSON 형식의 매핑 규칙 파일이 아닙니다.');
    }
    const loadedMapping = payload.mapping || payload; // 단순 {b: a} 형식도 허용

    if (headersB.length === 0) {
      throw new Error('먼저 타겟(B) 파일을 선택하고 [구조 분석]을 실행한 뒤 매핑 규칙을 불러오세요.');
    }

    let appliedCount = 0;
    let skippedCount = 0;
    Object.keys(loadedMapping).forEach((bCol) => {
      const select = mappingTableBody.querySelector(`select[data-bcolumn="${CSS.escape(bCol)}"]`);
      if (select) {
        const aCol = loadedMapping[bCol] || '';
        const optionExists = !aCol || Array.from(select.options).some((o) => o.value === aCol);
        select.value = optionExists ? aCol : '';
        mapping[bCol] = select.value;
        appliedCount += 1;
        if (aCol && !optionExists) {
          skippedCount += 1;
          log(`"${bCol}" 컬럼의 매핑 값("${aCol}")이 현재 원본(A) 파일에 없어 초기화했습니다.`, 'warn');
        }
      }
    });

    log(`매핑 규칙을 불러왔습니다: ${paths[0]} (적용 ${appliedCount}건, 불일치 ${skippedCount}건)`, 'ok');
  } catch (err) {
    notifyError('매핑 규칙 불러오기 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnConvert').addEventListener('click', async () => {
  try {
    if (!fileAPath || headersA.length === 0) {
      throw new Error('원본(A) 파일 분석이 필요합니다.');
    }
    if (headersB.length === 0) {
      throw new Error('타겟(B) 파일 분석이 필요합니다.');
    }

    const unmappedColumns = headersB.filter((b) => !mapping[b]);
    const mappedCount = headersB.length - unmappedColumns.length;
    if (mappedCount === 0) {
      throw new Error('매핑된 컬럼이 하나도 없습니다. 매핑 테이블을 확인하세요.');
    }
    if (unmappedColumns.length > 0) {
      log(`매핑되지 않은 컬럼 ${unmappedColumns.length}개는 빈 값으로 채워집니다: ${unmappedColumns.join(', ')}`, 'warn');
    }

    log(`데이터 변환을 시작합니다... (A 시트: ${sheetAName || '첫 번째 시트'})`);

    const workbook = readWorkbookSmart(fileAPath);
    const targetSheetName = sheetAName || workbook.SheetNames[0];
    const sheet = workbook.Sheets[targetSheetName];
    if (!sheet) throw new Error(`원본(A) 파일에서 "${targetSheetName}" 시트를 찾을 수 없습니다.`);
    const sourceRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (sourceRows.length === 0) {
      throw new Error('원본(A) 파일에 변환할 데이터 행이 없습니다.');
    }

    const convertedRows = sourceRows.map((row) => {
      const newRow = {};
      headersB.forEach((bCol) => {
        const aCol = mapping[bCol];
        newRow[bCol] = aCol ? (row[aCol] !== undefined ? row[aCol] : '') : '';
      });
      return newRow;
    });

    log(`총 ${convertedRows.length}행이 변환되었습니다.`, 'ok');

    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '변환된 엑셀 파일 저장',
      defaultPath: 'converted_result.xlsx',
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }],
    });
    if (!savePath) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }

    const newSheet = XLSX.utils.json_to_sheet(convertedRows, { header: headersB });
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Converted');
    const outputBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });
    safeWriteFileSync(savePath, outputBuffer);

    log(`변환된 파일이 저장되었습니다: ${savePath}`, 'ok');
    alert('✅ 변환이 완료되었습니다.');
  } catch (err) {
    notifyError('데이터 변환 중 오류가 발생했습니다', err);
  }
});

/* =========================================================================
   TAB 2. 메일머지 자동기입
   ========================================================================= */

let templatePath = null;
let templateVars = [];

const dropTemplate = document.getElementById('dropTemplate');
const fileNameTemplate = document.getElementById('fileNameTemplate');
const extractStatus = document.getElementById('extractStatus');
const varFormGrid = document.getElementById('varFormGrid');

bindDropzone(dropTemplate, fileNameTemplate, DOCX_FILTERS, (p) => {
  templatePath = p;
  log(`Word 템플릿 선택됨: ${p}`);
});

/** docx 파일 내부 word/document.xml 에서 {{변수명}} 패턴을 추출 */
function extractTemplateVariables(filePath) {
  const content = safeReadFileSync(filePath, 'binary');
  const zip = new PizZip(content);

  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('올바른 .docx 파일이 아닙니다. (word/document.xml 없음)');
  }
  const xml = docXmlFile.asText();

  // Word 는 텍스트를 여러 <w:t> 런(run)으로 쪼개는 경우가 많으므로,
  // XML 태그를 제거한 순수 텍스트에서 {{변수명}} 패턴을 탐색한다.
  const plainText = xml.replace(/<[^>]+>/g, '');

  const regex = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const found = new Set();
  let match;
  while ((match = regex.exec(plainText)) !== null) {
    const varName = match[1].trim();
    if (varName) found.add(varName);
  }

  return Array.from(found);
}

const varInputs = {}; // { varName: <input> }

function renderVarForm() {
  varFormGrid.innerHTML = '';
  Object.keys(varInputs).forEach((k) => delete varInputs[k]);

  if (templateVars.length === 0) {
    varFormGrid.innerHTML = '<div class="empty-hint">템플릿에서 {{변수명}} 형식의 항목을 찾지 못했습니다.</div>';
    return;
  }

  templateVars.forEach((varName) => {
    const wrap = document.createElement('div');
    wrap.className = 'var-field';

    const label = document.createElement('label');
    label.innerHTML = `<span class="tag">{{${varName}}}</span>`;
    wrap.appendChild(label);

    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = `${varName} 값을 입력하세요`;
    input.dataset.varName = varName;
    wrap.appendChild(input);

    varFormGrid.appendChild(wrap);
    varInputs[varName] = input;
  });
}

document.getElementById('btnExtractVars').addEventListener('click', () => {
  try {
    if (!templatePath) {
      throw new Error('Word 템플릿(.docx) 파일을 먼저 선택하세요.');
    }
    log('템플릿 변수 추출을 시작합니다...');
    templateVars = extractTemplateVariables(templatePath);

    if (templateVars.length === 0) {
      extractStatus.textContent = '변수 없음';
      extractStatus.className = 'pill warn';
      log('템플릿에서 {{변수명}} 패턴을 찾지 못했습니다.', 'warn');
    } else {
      extractStatus.textContent = `${templateVars.length}개 변수 발견`;
      extractStatus.className = 'pill ok';
      log(`변수 ${templateVars.length}개를 발견했습니다: ${templateVars.join(', ')}`, 'ok');
    }

    renderVarForm();
  } catch (err) {
    extractStatus.textContent = '추출 실패';
    extractStatus.className = 'pill warn';
    notifyError('변수 추출 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnGenerateDocx').addEventListener('click', async () => {
  try {
    if (!templatePath) {
      throw new Error('Word 템플릿(.docx) 파일을 먼저 선택하세요.');
    }
    if (templateVars.length === 0) {
      throw new Error('추출된 변수가 없습니다. 먼저 [변수 자동 추출]을 실행하세요.');
    }

    const data = {};
    const emptyVars = [];
    templateVars.forEach((varName) => {
      const input = varInputs[varName];
      data[varName] = input ? input.value : '';
      if (!data[varName]) emptyVars.push(varName);
    });
    if (emptyVars.length > 0) {
      log(`값이 입력되지 않은 변수 ${emptyVars.length}개는 빈 문자열로 치환됩니다: ${emptyVars.join(', ')}`, 'warn');
    }

    log('메일머지 문서 생성을 시작합니다...');

    const content = safeReadFileSync(templatePath, 'binary');
    const zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: '{{', end: '}}' },
    });

    try {
      doc.render(data);
    } catch (renderErr) {
      const details = (renderErr.properties && renderErr.properties.errors) || [];
      const detailMsg = details.map((e) => e.properties && e.properties.explanation).filter(Boolean).join('\n');
      throw new Error(detailMsg || renderErr.message || '템플릿 렌더링 중 오류가 발생했습니다.');
    }

    const outputBuffer = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });

    const baseName = path.basename(templatePath, path.extname(templatePath));
    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '메일머지 결과 저장',
      defaultPath: `${baseName}_결과.docx`,
      filters: DOCX_FILTERS,
    });
    if (!savePath) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }

    safeWriteFileSync(savePath, outputBuffer);
    log(`메일머지 문서가 저장되었습니다: ${savePath}`, 'ok');
    alert('✅ 문서 생성이 완료되었습니다.');
  } catch (err) {
    notifyError('문서 생성 중 오류가 발생했습니다', err);
  }
});

log('프로그램이 준비되었습니다.', 'ok');
