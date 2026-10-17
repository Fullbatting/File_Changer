// tab1.js
// TAB 1. 문서 양식 변환기: 엑셀/CSV 구조 분석, 컬럼 매핑(+프리셋), 검증/미리보기, 서식 적용 저장(ETL)

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const common = require('./common.js');

const {
  log,
  notifyError,
  friendlyErrorMessage,
  safeReadFileSync,
  safeWriteFileSync,
  confirmModal,
  readWorkbookSmart,
  DATA_FILE_FILTERS,
  JSON_FILTERS,
  bindDropzone,
  refreshSheetSelector,
  extractHeaders,
  findAutoMatch,
  xmlEscape,
  formatPreviewValue,
  getUserDataPath,
  onUserDataPathReady,
} = common;

let fileAPath = null;
let fileBPath = null;
let sheetAName = null;
let sheetBName = null;
let headerRowIndexA = 0; // 자동 감지된 헤더 행의 실제 시트 행 인덱스(0-based)
let headerRowIndexB = 0;
let headersA = [];
let headersB = [];
let mapping = {}; // { bColumn: aColumn|null }
let requiredColumns = new Set(); // 저장 전 검증에서 "필수"로 체크된 B 컬럼
let sampleRowsA = []; // 원본(A) 파일의 앞 2행 (매핑 테이블의 값 미리보기용)

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
const presetSelect = document.getElementById('presetSelect');

// 시트 선택 <select> 클릭이 드롭존의 파일선택 클릭으로 전파되지 않도록 차단
[sheetRowA, sheetRowB].forEach((row) => {
  row.addEventListener('click', (e) => e.stopPropagation());
});

/** 파일이 (재)선택되면 이전 분석 결과가 새 파일과 어긋나지 않도록 매핑 상태를 초기화 */
function resetAnalysisState() {
  headersA = [];
  headersB = [];
  headerRowIndexA = 0;
  headerRowIndexB = 0;
  mapping = {};
  requiredColumns = new Set();
  sampleRowsA = [];
  currentPresetId = null;
  renderMappingTable();
  if (presetSelect) presetSelect.value = '';
  analyzeStatus.textContent = '대기 중';
  analyzeStatus.className = 'pill warn';
}

/** 탭1 A/B 시트를 사용자가 직접 바꿨을 때: 분석 결과가 새 시트와 어긋나지 않도록 초기화 */
function onTab1SheetChanged(name) {
  resetAnalysisState();
  log(`시트가 "${name}"(으)로 변경되었습니다. [구조 분석]을 다시 실행하세요.`, 'warn');
}

bindDropzone(dropA, fileNameA, DATA_FILE_FILTERS, (p) => {
  fileAPath = p;
  log(`원본(A) 파일 선택됨: ${p}`);
  resetAnalysisState();
  refreshSheetSelector(
    p,
    sheetRowA,
    sheetSelectA,
    (name) => {
      sheetAName = name;
    },
    onTab1SheetChanged
  );
});

bindDropzone(dropB, fileNameB, DATA_FILE_FILTERS, (p) => {
  fileBPath = p;
  log(`타겟(B) 파일 선택됨: ${p}`);
  resetAnalysisState();
  refreshSheetSelector(
    p,
    sheetRowB,
    sheetSelectB,
    (name) => {
      sheetBName = name;
    },
    onTab1SheetChanged
  );
});

/** sampleRowsA에서 aCol 컬럼의 실제 값(앞 2행)을 뽑아 미리보기 텍스트로 만든다 */
function buildColumnPreview(aCol) {
  if (!aCol) return '';
  return sampleRowsA
    .map((row) => formatPreviewValue(row[aCol]))
    .filter((v) => v !== '')
    .join(' / ');
}

/**
 * 매핑 테이블을 그린다. 자동/수동으로 선택된 원본(A) 컬럼의 실제 값(앞 2행)을
 * "미리보기" 칸에 함께 보여줘, 변환을 실행하기 전에 매핑이 맞는지 바로 확인할
 * 수 있게 한다 (CSV/엑셀 임포트 도구들의 "Match Columns" 단계에서 흔히 쓰이는 패턴).
 */
function renderMappingTable() {
  mappingTableBody.innerHTML = '';
  mapping = {};
  requiredColumns = new Set();

  if (headersB.length === 0) {
    mappingTableBody.innerHTML = '<tr><td colspan="4" class="empty-hint">먼저 [구조 분석]을 실행하세요.</td></tr>';
    return;
  }

  const usedHeadersA = new Set();

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

    // 이름이 동일(완전 일치)/동의어 사전에 등록됨/편집거리상 비슷함 중 하나면 자동 매핑
    const tdPreview = document.createElement('td');
    tdPreview.className = 'mapping-preview-cell';

    const { match: autoMatch, reason } = findAutoMatch(bCol, headersA, usedHeadersA);
    if (autoMatch) {
      select.value = autoMatch;
      mapping[bCol] = autoMatch;
      usedHeadersA.add(autoMatch);
      if (reason === 'synonym') {
        log(`"${bCol}" 컬럼을 동의어 사전에 따라 "${autoMatch}" 컬럼에 자동 매핑했습니다.`);
      } else if (reason === 'fuzzy') {
        log(`"${bCol}" 컬럼을 이름이 비슷한 "${autoMatch}" 컬럼에 자동 매핑했습니다. 맞는지 확인하세요.`, 'warn');
      }
    } else {
      mapping[bCol] = '';
    }
    tdPreview.textContent = buildColumnPreview(mapping[bCol]);

    select.addEventListener('change', () => {
      mapping[bCol] = select.value;
      tdPreview.textContent = buildColumnPreview(select.value);
    });

    tdSelect.appendChild(select);
    tr.appendChild(tdSelect);
    tr.appendChild(tdPreview);

    const tdRequired = document.createElement('td');
    tdRequired.style.textAlign = 'center';
    const requiredCheckbox = document.createElement('input');
    requiredCheckbox.type = 'checkbox';
    requiredCheckbox.dataset.bcolumn = bCol;
    requiredCheckbox.addEventListener('change', () => {
      if (requiredCheckbox.checked) requiredColumns.add(bCol);
      else requiredColumns.delete(bCol);
    });
    tdRequired.appendChild(requiredCheckbox);
    tr.appendChild(tdRequired);

    mappingTableBody.appendChild(tr);
  });
}

/** requiredColumns 상태를 매핑 테이블의 "필수" 체크박스에도 반영(반영 안 된 나머지는 해제) */
function applyRequiredColumnsToTable(columns) {
  requiredColumns = new Set(columns || []);
  mappingTableBody.querySelectorAll('input[type="checkbox"][data-bcolumn]').forEach((checkbox) => {
    checkbox.checked = requiredColumns.has(checkbox.dataset.bcolumn);
  });
}

document.getElementById('btnAnalyze').addEventListener('click', () => {
  try {
    if (!fileAPath || !fileBPath) {
      throw new Error('원본(A) 파일과 타겟 양식(B) 파일을 모두 선택하세요.');
    }
    log(`구조 분석을 시작합니다... (A 시트: ${sheetAName || '첫 번째 시트'}, B 시트: ${sheetBName || '첫 번째 시트'})`);
    const resultA = extractHeaders(fileAPath, sheetAName);
    const resultB = extractHeaders(fileBPath, sheetBName);
    headersA = resultA.headers;
    headersB = resultB.headers;
    headerRowIndexA = resultA.headerRowIndex;
    headerRowIndexB = resultB.headerRowIndex;
    if (headerRowIndexA > 0) {
      log(`원본(A) 파일: 제목/빈 행으로 보이는 ${headerRowIndexA}개 행을 건너뛰고 ${headerRowIndexA + 1}행을 헤더로 인식했습니다.`, 'warn');
    }
    if (headerRowIndexB > 0) {
      log(`타겟(B) 파일: 제목/빈 행으로 보이는 ${headerRowIndexB}개 행을 건너뛰고 ${headerRowIndexB + 1}행을 헤더로 인식했습니다.`, 'warn');
    }
    log(`원본(A) 컬럼 ${headersA.length}개, 타겟(B) 컬럼 ${headersB.length}개 발견`, 'ok');

    // 매핑 테이블에 값 미리보기를 보여주기 위해 원본(A)의 앞 2행을 함께 읽어둔다.
    const workbookA = readWorkbookSmart(fileAPath);
    const sheetA = workbookA.Sheets[sheetAName || workbookA.SheetNames[0]];
    const previewRowsRaw = XLSX.utils.sheet_to_json(sheetA, { defval: '', range: headerRowIndexA });
    sampleRowsA = previewRowsRaw.slice(0, 2);

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
      requiredColumns: Array.from(requiredColumns),
    };
    safeWriteFileSync(savePath, JSON.stringify(payload, null, 2));
    log(`매핑 규칙이 저장되었습니다: ${savePath}`, 'ok');
  } catch (err) {
    notifyError('매핑 규칙 저장 중 오류가 발생했습니다', err);
  }
});

/**
 * loadedMapping({ bColumn: aColumn }) 을 현재 매핑 테이블에 적용한다.
 * 현재 원본(A) 파일에 없는 컬럼 값은 초기화하고 경고 로그를 남긴다.
 * 매핑 규칙 파일 불러오기 / 프리셋 적용에서 공통으로 사용.
 */
function applyMappingToTable(loadedMapping) {
  let appliedCount = 0;
  let skippedCount = 0;
  Object.keys(loadedMapping).forEach((bCol) => {
    const select = mappingTableBody.querySelector(`select[data-bcolumn="${CSS.escape(bCol)}"]`);
    if (select) {
      const aCol = loadedMapping[bCol] || '';
      const optionExists = !aCol || Array.from(select.options).some((o) => o.value === aCol);
      select.value = optionExists ? aCol : '';
      mapping[bCol] = select.value;
      const previewCell = select.closest('tr').querySelector('.mapping-preview-cell');
      if (previewCell) previewCell.textContent = buildColumnPreview(select.value);
      appliedCount += 1;
      if (aCol && !optionExists) {
        skippedCount += 1;
        log(`"${bCol}" 컬럼의 매핑 값("${aCol}")이 현재 원본(A) 파일에 없어 초기화했습니다.`, 'warn');
      }
    }
  });
  return { appliedCount, skippedCount };
}

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

    const { appliedCount, skippedCount } = applyMappingToTable(loadedMapping);
    if (Array.isArray(payload.requiredColumns)) {
      applyRequiredColumnsToTable(payload.requiredColumns.filter((c) => headersB.includes(c)));
    }
    currentPresetId = null;
    if (presetSelect) presetSelect.value = '';
    log(`매핑 규칙을 불러왔습니다: ${paths[0]} (적용 ${appliedCount}건, 불일치 ${skippedCount}건)`, 'ok');
  } catch (err) {
    notifyError('매핑 규칙 불러오기 중 오류가 발생했습니다', err);
  }
});

/* -------------------------------------------------------------------------
 * 매핑 프리셋 (이름 붙여 이 PC에 저장, 수정/삭제 가능)
 * 파일로 내보내는 매핑 규칙(.json)과 달리, 앱 내부(사용자 데이터 폴더)에
 * 이름과 함께 저장되어 다음 실행 시에도 드롭다운에서 바로 선택할 수 있다.
 * ------------------------------------------------------------------------- */

let presets = [];
let currentPresetId = null;
const PRESETS_FILE_NAME = 'mapping-presets.json';

function getPresetsFilePath() {
  return path.join(getUserDataPath(), PRESETS_FILE_NAME);
}

function generatePresetId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function loadPresetsFromDisk() {
  try {
    if (!fs.existsSync(getPresetsFilePath())) {
      presets = [];
      return;
    }
    const raw = safeReadFileSync(getPresetsFilePath(), 'utf-8');
    const parsed = JSON.parse(raw);
    presets = Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    presets = [];
    log(`저장된 프리셋을 불러오는 중 오류가 발생하여 목록을 초기화합니다: ${friendlyErrorMessage(err)}`, 'warn');
  }
  renderPresetSelect();
}

function savePresetsToDisk() {
  safeWriteFileSync(getPresetsFilePath(), JSON.stringify(presets, null, 2));
}

function renderPresetSelect() {
  if (!presetSelect) return;
  const sorted = [...presets].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
  presetSelect.innerHTML = '<option value="">-- 새 프리셋 --</option>';
  sorted.forEach((p) => {
    const opt = document.createElement('option');
    opt.value = p.id;
    opt.textContent = p.name;
    presetSelect.appendChild(opt);
  });
  presetSelect.value = currentPresetId || '';
}

// 프리셋 이름 입력 모달
const presetModal = document.getElementById('presetNameModal');
const presetNameInput = document.getElementById('presetNameInput');
const presetModalError = document.getElementById('presetModalError');
const presetModalSaveBtn = document.getElementById('presetModalSave');
const presetModalCancelBtn = document.getElementById('presetModalCancel');

/** 프리셋 이름 입력 모달을 띄우고, 저장된 이름(취소 시 null)을 Promise로 반환 */
function showPresetNameModal(defaultValue) {
  return new Promise((resolve) => {
    presetNameInput.value = defaultValue || '';
    presetModalError.hidden = true;
    presetModal.hidden = false;
    presetNameInput.focus();
    presetNameInput.select();

    const cleanup = () => {
      presetModal.hidden = true;
      presetModalSaveBtn.onclick = null;
      presetModalCancelBtn.onclick = null;
      presetNameInput.onkeydown = null;
      presetModal.onclick = null;
    };

    presetModalSaveBtn.onclick = () => {
      const name = presetNameInput.value.trim();
      if (!name) {
        presetModalError.textContent = '프리셋 이름을 입력하세요.';
        presetModalError.hidden = false;
        return;
      }
      cleanup();
      resolve(name);
    };

    presetModalCancelBtn.onclick = () => {
      cleanup();
      resolve(null);
    };

    presetNameInput.onkeydown = (e) => {
      if (e.key === 'Enter') presetModalSaveBtn.click();
      if (e.key === 'Escape') presetModalCancelBtn.click();
    };

    // 배경(오버레이) 클릭 시 취소
    presetModal.onclick = (e) => {
      if (e.target === presetModal) presetModalCancelBtn.click();
    };
  });
}

if (presetSelect) {
  presetSelect.addEventListener('change', (e) => {
    const id = e.target.value;
    if (!id) {
      currentPresetId = null;
      return;
    }
    const preset = presets.find((p) => p.id === id);
    if (!preset) return;

    if (headersB.length === 0) {
      alert('먼저 원본(A)/타겟(B) 파일을 선택하고 [구조 분석]을 실행하세요.');
      e.target.value = currentPresetId || '';
      return;
    }

    const { appliedCount, skippedCount } = applyMappingToTable(preset.mapping);
    if (Array.isArray(preset.requiredColumns)) {
      applyRequiredColumnsToTable(preset.requiredColumns.filter((c) => headersB.includes(c)));
    }
    currentPresetId = id;
    log(`프리셋 "${preset.name}"을(를) 적용했습니다. (적용 ${appliedCount}건, 불일치 ${skippedCount}건)`, 'ok');
  });
}

document.getElementById('btnSavePreset').addEventListener('click', async () => {
  try {
    if (headersB.length === 0) throw new Error('저장할 매핑이 없습니다. 먼저 구조 분석을 실행하세요.');

    const current = presets.find((p) => p.id === currentPresetId);
    const name = await showPresetNameModal(current ? current.name : '');
    if (name === null) return; // 취소됨

    const existing = presets.find((p) => p.name === name);
    if (existing && existing.id !== currentPresetId) {
      const overwrite = confirm(`이미 "${name}" 프리셋이 있습니다. 덮어쓸까요?`);
      if (!overwrite) return;
    }

    const targetId = existing ? existing.id : currentPresetId || generatePresetId();
    const now = new Date().toISOString();
    const existingRecord = presets.find((p) => p.id === targetId);

    const preset = {
      id: targetId,
      name,
      sourceHeaders: headersA,
      targetHeaders: headersB,
      mapping: { ...mapping },
      requiredColumns: Array.from(requiredColumns),
      createdAt: existingRecord ? existingRecord.createdAt : now,
      updatedAt: now,
    };

    const idx = presets.findIndex((p) => p.id === targetId);
    if (idx >= 0) presets[idx] = preset;
    else presets.push(preset);

    savePresetsToDisk();
    currentPresetId = targetId;
    renderPresetSelect();
    log(`프리셋 "${name}"(으)로 저장되었습니다.`, 'ok');
  } catch (err) {
    notifyError('프리셋 저장 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnDeletePreset').addEventListener('click', () => {
  try {
    if (!currentPresetId) throw new Error('삭제할 프리셋을 선택하세요.');
    const preset = presets.find((p) => p.id === currentPresetId);
    if (!preset) throw new Error('선택된 프리셋을 찾을 수 없습니다.');

    const ok = confirm(`"${preset.name}" 프리셋을 삭제할까요? 이 동작은 되돌릴 수 없습니다.`);
    if (!ok) return;

    presets = presets.filter((p) => p.id !== currentPresetId);
    savePresetsToDisk();
    currentPresetId = null;
    renderPresetSelect();
    log(`프리셋 "${preset.name}"이(가) 삭제되었습니다.`, 'ok');
  } catch (err) {
    notifyError('프리셋 삭제 중 오류가 발생했습니다', err);
  }
});

// 사용자 데이터 폴더 경로가 준비되면(공통 모듈이 비동기로 전달) 저장된 프리셋을 불러온다.
onUserDataPathReady(() => loadPresetsFromDisk());

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

/**
 * 값이 "숫자로 바꿔도 안전한" 문자열인지 판정한다. 단순히 숫자처럼 보이는지가
 * 아니라, Number로 바꾼 뒤 다시 문자열화했을 때 원래 값과 완전히 같아야
 * 통과한다 — 우편번호·사번처럼 앞자리가 0인 코드성 값은 숫자로 바꾸면
 * 자릿수가 사라지므로(예: "03187" → 3187) 이 조건에서 자동으로 걸러진다.
 */
function isSafeNumericString(v) {
  if (typeof v === 'number') return Number.isFinite(v);
  if (typeof v !== 'string') return false;
  const trimmed = v.trim();
  if (trimmed === '' || !/^-?\d+(\.\d+)?$/.test(trimmed)) return false;
  return String(Number(trimmed)) === trimmed;
}

/**
 * 변환된 데이터를 저장하기 전에 검증한다.
 * - 필수로 지정된 B 컬럼에 빈 값이 있는 행의 수와 예시를 모은다.
 * - 전체 컬럼 값이 완전히 똑같은 행(중복 행)의 수와 예시를 모은다.
 * (react-spreadsheet-import 등 CSV 임포트 도구들의 "Validate Data" 단계를
 * 참고했으며, 문제가 있어도 저장을 막지는 않고 사용자가 확인 후 진행/취소를
 * 선택하게 한다 — 그 도구들의 allowInvalidSubmit 기본 동작과 동일한 취지.)
 */
function validateConvertedRows(convertedRows, headersB, requiredCols) {
  const missingByColumn = {};
  requiredCols.forEach((col) => {
    missingByColumn[col] = 0;
  });

  const seen = new Map(); // 행 내용(JSON) -> 처음 나온 행 번호(1-based)
  let duplicateCount = 0;
  const duplicateExamples = [];

  convertedRows.forEach((row, idx) => {
    requiredCols.forEach((col) => {
      const v = row[col];
      if (v === '' || v === null || v === undefined) missingByColumn[col] += 1;
    });

    const key = headersB.map((h) => row[h]).join('\u0001');
    if (seen.has(key)) {
      duplicateCount += 1;
      if (duplicateExamples.length < 5) {
        duplicateExamples.push({ row: idx + 1, duplicateOf: seen.get(key) + 1 });
      }
    } else {
      seen.set(key, idx);
    }
  });

  return { missingByColumn, duplicateCount, duplicateExamples };
}

// 변환 결과 확인(검증 + 미리보기) 모달
const convertReviewModal = document.getElementById('convertReviewModal');
const convertReviewSummary = document.getElementById('convertReviewSummary');
const convertReviewTableHead = document.getElementById('convertReviewTableHead');
const convertReviewTableBody = document.getElementById('convertReviewTableBody');
const convertReviewProceedBtn = document.getElementById('convertReviewProceed');
const convertReviewCancelBtn = document.getElementById('convertReviewCancel');

const CONVERT_PREVIEW_ROW_LIMIT = 5;

/**
 * 변환 결과 확인 모달을 띄운다. 검증 요약 + 앞부분 미리보기를 보여주고,
 * "그대로 저장"(true) / "취소하고 수정"(false)을 Promise로 반환한다.
 */
function showConvertReviewModal(headersB, convertedRows, validation) {
  const lines = [];
  lines.push(`<div class="review-line ok">총 <b>${convertedRows.length}</b>행이 변환되었습니다.</div>`);

  const missingEntries = Object.entries(validation.missingByColumn).filter(([, count]) => count > 0);
  if (missingEntries.length > 0) {
    const detail = missingEntries.map(([col, count]) => `"${xmlEscape(col)}" ${count}건`).join(', ');
    lines.push(`<div class="review-line warn">⚠ 필수 컬럼 값이 비어 있는 행: ${detail}</div>`);
  } else if (Object.keys(validation.missingByColumn).length > 0) {
    lines.push(`<div class="review-line ok">✓ 필수 컬럼에 빈 값이 없습니다.</div>`);
  }

  if (validation.duplicateCount > 0) {
    const examples = validation.duplicateExamples.map((e) => `${e.row}행(${e.duplicateOf}행과 동일)`).join(', ');
    lines.push(`<div class="review-line warn">⚠ 모든 컬럼 값이 동일한 중복 행 ${validation.duplicateCount}건: ${examples}${validation.duplicateCount > validation.duplicateExamples.length ? ' 등' : ''}</div>`);
  } else {
    lines.push(`<div class="review-line ok">✓ 완전히 중복된 행이 없습니다.</div>`);
  }

  convertReviewSummary.innerHTML = lines.join('');

  convertReviewTableHead.innerHTML = headersB.map((h) => `<th>${xmlEscape(h)}</th>`).join('');
  const previewRows = convertedRows.slice(0, CONVERT_PREVIEW_ROW_LIMIT);
  convertReviewTableBody.innerHTML = previewRows
    .map(
      (row) =>
        `<tr>${headersB
          .map((h) => `<td>${row[h] === '' || row[h] === null || row[h] === undefined ? '' : xmlEscape(String(row[h]))}</td>`)
          .join('')}</tr>`
    )
    .join('');
  if (convertedRows.length > previewRows.length) {
    convertReviewTableBody.innerHTML += `<tr><td colspan="${headersB.length}" class="empty-hint">... 외 ${convertedRows.length - previewRows.length}행</td></tr>`;
  }

  return confirmModal(convertReviewModal, convertReviewProceedBtn, convertReviewCancelBtn);
}

/**
 * [Load] 변환된 데이터를 서식이 적용된 엑셀 파일로 내보낸다 (exceljs 사용).
 * - 헤더 행: 배경색 + 흰 글자 + 굵게 + 테두리, 첫 행 고정(스크롤해도 헤더 유지)
 * - 열 너비: 헤더/값의 실제 표시 폭에 맞춰 자동 조정 → "###" 잘림 방지
 * - 모든 값이 "안전한 숫자"인 열만 실제 숫자로 저장하고 천단위 구분(#,##0)
 *   서식과 우측 정렬을 적용. 그 외에는 원래 문자열 그대로 좌측 정렬 유지.
 */
async function buildStyledExcelBuffer(headers, rows, sheetName) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName || 'Converted');

  worksheet.addRow(headers);
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      left: { style: 'thin', color: { argb: 'FFB0B0B0' } },
      right: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    };
  });

  // 열별로 "빈 값을 제외한 모든 값이 안전한 숫자인지" 판정
  const columnIsNumeric = headers.map((h) => {
    let sawAny = false;
    for (const row of rows) {
      const v = row[h];
      if (v === '' || v === null || v === undefined) continue;
      sawAny = true;
      if (!isSafeNumericString(v)) return false;
    }
    return sawAny;
  });
  const columnHasDecimal = headers.map(
    (h, i) => columnIsNumeric[i] && rows.some((row) => String(row[h]).includes('.'))
  );

  rows.forEach((row) => {
    const values = headers.map((h, i) => {
      const raw = row[h];
      if (raw === '' || raw === null || raw === undefined) return '';
      return columnIsNumeric[i] ? Number(raw) : raw;
    });
    const excelRow = worksheet.addRow(values);
    excelRow.eachCell((cell, colNumber) => {
      const i = colNumber - 1;
      if (columnIsNumeric[i]) {
        cell.numFmt = columnHasDecimal[i] ? '#,##0.######' : '#,##0';
        cell.alignment = { horizontal: 'right' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    });
  });

  headers.forEach((h, i) => {
    let maxWidth = visualWidth(h);
    rows.forEach((row) => {
      const v = row[h];
      if (v === '' || v === null || v === undefined) return;
      maxWidth = Math.max(maxWidth, visualWidth(String(v)));
    });
    worksheet.getColumn(i + 1).width = Math.min(40, Math.max(8, Math.ceil(maxWidth) + 2));
  });

  return workbook.xlsx.writeBuffer();
}

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

    // [Extract] 원본(A) 파일에서 원시 데이터를 읽어온다 (자동 감지된 헤더 행 다음부터)
    log(`데이터 변환을 시작합니다... (A 시트: ${sheetAName || '첫 번째 시트'})`);

    const workbook = readWorkbookSmart(fileAPath);
    const targetSheetName = sheetAName || workbook.SheetNames[0];
    const sheet = workbook.Sheets[targetSheetName];
    if (!sheet) throw new Error(`원본(A) 파일에서 "${targetSheetName}" 시트를 찾을 수 없습니다.`);
    const sourceRows = XLSX.utils.sheet_to_json(sheet, { defval: '', range: headerRowIndexA });

    if (sourceRows.length === 0) {
      throw new Error('원본(A) 파일에 변환할 데이터 행이 없습니다.');
    }

    // [Transform] 매핑 규칙(코드와 분리되어 프리셋으로 저장/불러오기 가능)에 따라
    // A 양식의 각 행을 B 양식의 컬럼 구조로 재구성한다
    const convertedRows = sourceRows.map((row) => {
      const newRow = {};
      headersB.forEach((bCol) => {
        const aCol = mapping[bCol];
        newRow[bCol] = aCol ? (row[aCol] !== undefined ? row[aCol] : '') : '';
      });
      return newRow;
    });

    log(`총 ${convertedRows.length}행이 변환되었습니다.`, 'ok');

    // 저장 전 검증 + 미리보기 확인 (필수 컬럼 빈 값 / 완전 중복 행)
    const validation = validateConvertedRows(convertedRows, headersB, Array.from(requiredColumns));
    const missingTotal = Object.values(validation.missingByColumn).reduce((sum, n) => sum + n, 0);
    if (missingTotal > 0) {
      log(`필수 컬럼 값이 비어 있는 행이 있습니다 (총 ${missingTotal}건). 확인 모달에서 자세한 내용을 볼 수 있습니다.`, 'warn');
    }
    if (validation.duplicateCount > 0) {
      log(`모든 컬럼 값이 동일한 중복 행이 ${validation.duplicateCount}건 있습니다.`, 'warn');
    }

    const proceed = await showConvertReviewModal(headersB, convertedRows, validation);
    if (!proceed) {
      log('사용자가 확인 단계에서 취소했습니다. 매핑을 수정한 뒤 다시 시도하세요.', 'warn');
      return;
    }

    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '변환된 엑셀 파일 저장',
      defaultPath: 'converted_result.xlsx',
      filters: [{ name: 'Excel File', extensions: ['xlsx'] }],
    });
    if (!savePath) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }

    // [Load] 헤더 서식·열 너비·숫자 서식을 적용하여 완성된 엑셀 파일로 저장한다
    const outputBuffer = await buildStyledExcelBuffer(headersB, convertedRows, sheetBName);
    safeWriteFileSync(savePath, outputBuffer);

    log(`변환된 파일이 저장되었습니다: ${savePath}`, 'ok');
    alert('✅ 변환이 완료되었습니다.');
  } catch (err) {
    notifyError('데이터 변환 중 오류가 발생했습니다', err);
  }
});
