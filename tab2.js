// tab2.js
// 메일머지 자동기입 화면: 템플릿 고르기 → 변수 추출 → 직접 입력(1건) 또는 데이터 파일로 일괄 생성.
// 템플릿 찾기·바꾸기·파일 이름 규칙은 merge-template.js, 데이터 파일 읽기는 data-file.js(common.js가 감쌈)에 있고
// 여기는 화면 상태와 버튼만 다룬다.

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const { friendlyErrorMessage, safeWriteFileSync } = require('./file-utils.js');
const merge = require('./merge-template.js');
const { log, notifyError, setPill, confirmModal, bindDropzone, refreshSheetSelector, readSheetTable, describeRecognition, autoMatchAll, formatPreviewValue, DATA_FILE_FILTERS } = require('./common.js');

// 문서 계열: 본문 형식(docx/hwpx)과 그 서식 파일(dotx/hwpt)은 구조가 같아 같은 엔진으로 처리한다
const TEMPLATE_FILTERS = [
  { name: 'Word / 한글(HWPX) 문서', extensions: ['docx', 'dotx', 'hwpx', 'hwpt'] },
  { name: 'All Files', extensions: ['*'] },
];
const SAVE_FILTERS = {
  docx: [{ name: 'Word Document', extensions: ['docx'] }, { name: 'All Files', extensions: ['*'] }],
  hwpx: [{ name: '한글(HWPX) 문서', extensions: ['hwpx'] }, { name: 'All Files', extensions: ['*'] }],
};
const NO_TEMPLATE = 'Word/한글(.docx, .hwpx) 템플릿 파일을 먼저 선택하세요.';
const NO_VARIABLES = '추출된 변수가 없습니다. 먼저 [변수 자동 추출]을 실행하세요.';
const MAPPING_PREVIEW_ROWS = 2; // 매핑 표 미리보기에 보일 앞 행 수
const REVIEW_PREVIEW_ROWS = 5; // 일괄 생성 확인 창에 보일 앞 행 수
const YIELD_EVERY = 20; // 수백~수천 행을 만들 때 화면이 멈춰 보이지 않도록 이만큼마다 한 번 쉰다
const PROGRESS_LOG_EVERY = 100;

const state = {
  templatePath: null,
  variables: [], // 템플릿에서 찾은 {{변수}}
  mode: 'manual', // 'manual'(1건 직접 입력) | 'batch'(데이터 파일로 일괄 생성)
  data: { path: null, sheet: null, headers: [], rows: [] },
  mapping: {}, // { 템플릿 변수: 데이터 열 }
  filenameColumn: '',
};

const $ = (id) => document.getElementById(id);
const el = {
  extractStatus: $('extractStatus'),
  varFormGrid: $('varFormGrid'),
  manualCard: $('manualInputCard'),
  batchCard: $('batchInputCard'),
  btnGenerate: $('btnGenerateDocx'),
  btnGenerateBatch: $('btnGenerateBatch'),
  generateDesc: $('generateCardDesc'),
  dataStatus: $('mailMergeDataStatus'),
  sheetRow: $('sheetRowMailMerge'),
  sheetSelect: $('sheetSelectMailMerge'),
  mappingBody: $('mailMergeMappingTableBody'),
  filenameSelect: $('filenameColumnSelect'),
  review: {
    modal: $('mailMergeReviewModal'),
    summary: $('mailMergeReviewSummary'),
    head: $('mailMergeReviewTableHead'),
    body: $('mailMergeReviewTableBody'),
    proceed: $('mailMergeReviewProceed'),
    cancel: $('mailMergeReviewCancel'),
  },
};

/** 화면에 넣을 글자(HTML 특수문자 바꿈) */
const escapeHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** <select>에 첫 항목(빈 값)과 목록을 채운다 */
function fillSelect(select, emptyLabel, values) {
  select.innerHTML = '';
  select.appendChild(new Option(emptyLabel, ''));
  values.forEach((v) => select.appendChild(new Option(v, v)));
}

/* ---------- 1. 템플릿 ---------- */

bindDropzone($('dropTemplate'), $('fileNameTemplate'), TEMPLATE_FILTERS, (p) => {
  state.templatePath = p;
  log(`템플릿 선택됨: ${p}`);
});

$('btnExtractVars').addEventListener('click', () => {
  try {
    if (!state.templatePath) throw new Error(NO_TEMPLATE);
    const format = merge.templateFormat(state.templatePath);
    log(`템플릿 변수 추출을 시작합니다... (형식: ${format.toUpperCase()})`);
    state.variables = merge.openTemplate(state.templatePath).variables;
    if (state.variables.length === 0) {
      setPill(el.extractStatus, '변수 없음', false);
      log('템플릿에서 {{변수명}} 패턴을 찾지 못했습니다.', 'warn');
    } else {
      setPill(el.extractStatus, `${state.variables.length}개 변수 발견`, true);
      log(`변수 ${state.variables.length}개를 발견했습니다: ${state.variables.join(', ')}`, 'ok');
    }
    renderVarForm();
    renderMappingTable();
  } catch (err) {
    setPill(el.extractStatus, '추출 실패', false);
    notifyError('변수 추출 중 오류가 발생했습니다', err);
  }
});

/* ---------- 2. 입력 방식 ---------- */

document.querySelectorAll('input[name="mmMode"]').forEach((radio) => {
  radio.addEventListener('change', (e) => {
    state.mode = e.target.value;
    const isBatch = state.mode === 'batch';
    el.manualCard.hidden = isBatch;
    el.batchCard.hidden = !isBatch;
    el.btnGenerate.hidden = isBatch;
    el.btnGenerateBatch.hidden = !isBatch;
    el.generateDesc.textContent = isBatch ? '데이터 파일의 각 행마다 문서를 생성하여 선택한 폴더에 저장합니다.' : '입력한 값으로 템플릿 변수를 치환하여 결과 문서(.docx 또는 .hwpx)를 생성합니다.';
  });
});

/* ---------- 3-1. 직접 입력(1건) ---------- */

const varInputs = new Map(); // 변수 -> <input>

function renderVarForm() {
  el.varFormGrid.innerHTML = '';
  varInputs.clear();
  if (state.variables.length === 0) {
    el.varFormGrid.innerHTML = '<div class="empty-hint">템플릿에서 {{변수명}} 형식의 항목을 찾지 못했습니다.</div>';
    return;
  }
  state.variables.forEach((name) => {
    const wrap = document.createElement('div');
    wrap.className = 'var-field';
    const label = document.createElement('label');
    const tag = document.createElement('span');
    tag.className = 'tag';
    tag.textContent = `{{${name}}}`;
    label.appendChild(tag);
    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = `${name} 값을 입력하세요`;
    input.dataset.varName = name;
    wrap.append(label, input);
    el.varFormGrid.appendChild(wrap);
    varInputs.set(name, input);
  });
}

el.btnGenerate.addEventListener('click', async () => {
  try {
    if (!state.templatePath) throw new Error(NO_TEMPLATE);
    if (state.variables.length === 0) throw new Error(NO_VARIABLES);
    const format = merge.templateFormat(state.templatePath);
    const data = {};
    state.variables.forEach((name) => {
      data[name] = varInputs.has(name) ? varInputs.get(name).value : '';
    });
    const empty = state.variables.filter((name) => !data[name]);
    if (empty.length > 0) log(`값이 입력되지 않은 변수 ${empty.length}개는 빈 문자열로 치환됩니다: ${empty.join(', ')}`, 'warn');

    log(`메일머지 문서 생성을 시작합니다... (형식: ${format.toUpperCase()})`);
    const output = merge.openTemplate(state.templatePath).render(data);
    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '메일머지 결과 저장',
      defaultPath: `${path.basename(state.templatePath, path.extname(state.templatePath))}_결과.${format}`,
      filters: SAVE_FILTERS[format],
    });
    if (!savePath) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }
    safeWriteFileSync(savePath, output);
    log(`메일머지 문서가 저장되었습니다: ${savePath}`, 'ok');
    alert('✅ 문서 생성이 완료되었습니다.');
  } catch (err) {
    notifyError('문서 생성 중 오류가 발생했습니다', err);
  }
});

/* ---------- 3-2. 데이터 파일과 열 매핑 ---------- */

el.sheetRow.addEventListener('click', (e) => e.stopPropagation());

/** 데이터 파일·시트를 (다시) 고르면 불러온 행·매핑을 비운다(새 시트와 어긋나지 않게) */
function resetData() {
  state.data.headers = [];
  state.data.rows = [];
  renderMappingTable();
  setPill(el.dataStatus, '대기 중', false);
}

bindDropzone($('dropMailMergeData'), $('fileNameMailMergeData'), DATA_FILE_FILTERS, (p) => {
  state.data.path = p;
  log(`메일머지 데이터 파일 선택됨: ${p}`);
  resetData();
  refreshSheetSelector(
    p,
    el.sheetRow,
    el.sheetSelect,
    (name) => {
      state.data.sheet = name;
    },
    (name) => {
      resetData();
      log(`시트가 "${name}"(으)로 변경되었습니다. [데이터 불러오기]를 다시 실행하세요.`, 'warn');
    }
  );
});

$('btnLoadMailMergeData').addEventListener('click', () => {
  try {
    if (!state.data.path) throw new Error('데이터 파일을 먼저 선택하세요.');
    if (state.variables.length === 0) throw new Error('먼저 템플릿을 선택하고 [변수 자동 추출]을 실행하세요.');
    log(`메일머지 데이터를 불러옵니다... (시트: ${state.data.sheet || '첫 번째 시트'})`);
    // displayText: 문서에는 엑셀 화면에 보이는 그대로(1,500,000 / 15% / ₩1,000) 들어가야 한다
    const table = readSheetTable(state.data.path, state.data.sheet, { displayText: true });
    state.data.headers = table.headers;
    describeRecognition(table.report, table.fileInfo).forEach((item) => log(`데이터 파일 인식: ${item.text}`, item.warn ? 'warn' : 'info'));
    const renamed = table.headers.filter((k, i) => k !== table.labels[i]);
    if (renamed.length > 0) log(`이름이 같은 컬럼이 있어 ${renamed.join(', ')}(으)로 구분했습니다.`, 'warn');
    state.data.rows = table.rows;
    if (state.data.rows.length === 0) throw new Error('데이터 파일에 변환할 행이 없습니다.');

    renderMappingTable();
    renderFilenameOptions();
    setPill(el.dataStatus, `${state.data.rows.length}행 불러옴`, true);
    log(`데이터 ${state.data.rows.length}행, 컬럼 ${state.data.headers.length}개를 불러왔습니다.`, 'ok');
  } catch (err) {
    setPill(el.dataStatus, '불러오기 실패', false);
    notifyError('메일머지 데이터를 불러오는 중 오류가 발생했습니다', err);
  }
});

/** 열의 실제 값(앞 몇 행) — 매핑이 맞는지 일괄 생성 전에 바로 보이게 */
const columnPreview = (column) =>
  column
    ? state.data.rows
        .slice(0, MAPPING_PREVIEW_ROWS)
        .map((row) => formatPreviewValue(row[column]))
        .filter((v) => v !== '')
        .join(' / ')
    : '';

/** 템플릿 변수 ← 데이터 열 매핑 표. 자동 매핑(같은 이름 → 동의어 → 비슷한 이름)을 먼저 채운다 */
function renderMappingTable() {
  el.mappingBody.innerHTML = '';
  state.mapping = {};
  if (state.variables.length === 0 || state.data.headers.length === 0) {
    el.mappingBody.innerHTML = '<tr><td colspan="3" class="empty-hint">먼저 [변수 자동 추출]과 [데이터 불러오기]를 실행하세요.</td></tr>';
    return;
  }
  const auto = autoMatchAll(state.variables, state.data.headers);
  state.variables.forEach((name) => {
    const { match, reason } = auto[name];
    state.mapping[name] = match || '';
    if (reason === 'synonym') log(`변수 "${name}"을 동의어 사전에 따라 데이터 컬럼 "${match}"에 자동 매핑했습니다.`);
    else if (reason === 'fuzzy') log(`변수 "${name}"을 이름이 비슷한 데이터 컬럼 "${match}"에 자동 매핑했습니다. 맞는지 확인하세요.`, 'warn');

    const tr = document.createElement('tr');
    const tdVar = document.createElement('td');
    tdVar.textContent = `{{${name}}}`;
    const select = document.createElement('select');
    select.dataset.templateVar = name;
    fillSelect(select, '-- 선택 안함 (빈 값) --', state.data.headers);
    select.value = state.mapping[name];
    const tdSelect = document.createElement('td');
    tdSelect.appendChild(select);
    const tdPreview = document.createElement('td');
    tdPreview.className = 'mapping-preview-cell';
    tdPreview.textContent = columnPreview(state.mapping[name]);
    select.addEventListener('change', () => {
      state.mapping[name] = select.value;
      tdPreview.textContent = columnPreview(select.value);
    });
    tr.append(tdVar, tdSelect, tdPreview);
    el.mappingBody.appendChild(tr);
  });
}

/** 결과 파일 이름 기준 열. "이름"·"성명" 류 열이 있으면 먼저 골라 둔다 */
function renderFilenameOptions() {
  fillSelect(el.filenameSelect, '-- 행 번호 사용 --', state.data.headers);
  state.filenameColumn = merge.defaultFilenameColumn(state.data.headers);
  el.filenameSelect.value = state.filenameColumn;
  el.filenameSelect.onchange = () => {
    state.filenameColumn = el.filenameSelect.value;
  };
}

/* ---------- 4. 일괄 생성 ---------- */

/** 일괄 생성 확인 창(빈 값·파일 이름 검사 + 앞 몇 행 미리보기). "그대로 생성"이면 true */
function showReview(filenames) {
  const { rows } = state.data;
  const lines = [`<div class="review-line ok">데이터 ${rows.length}행 → 문서 <b>${rows.length}</b>개를 생성합니다.</div>`];
  const missing = Object.entries(merge.missingValueCounts(rows, state.variables, state.mapping)).filter(([, n]) => n > 0);
  lines.push(
    missing.length > 0
      ? `<div class="review-line warn">⚠ 값이 비어 있는 변수: ${missing.map(([v, n]) => `"${escapeHtml(v)}" ${n}건`).join(', ')}</div>`
      : '<div class="review-line ok">✓ 모든 변수에 값이 채워져 있습니다.</div>'
  );
  const unnamed = merge.unnamedRowCount(rows, state.filenameColumn);
  if (unnamed > 0) lines.push(`<div class="review-line warn">⚠ 파일명 기준 컬럼 값이 비어 있어 행 번호로 이름 지어지는 행: ${unnamed}건</div>`);
  el.review.summary.innerHTML = lines.join('');

  el.review.head.innerHTML = ['파일명', ...state.variables].map((h) => `<th>${escapeHtml(h)}</th>`).join('');
  const shown = rows.slice(0, REVIEW_PREVIEW_ROWS).map((row, i) => {
    const cells = [filenames[i], ...state.variables.map((v) => formatPreviewValue(state.mapping[v] ? row[state.mapping[v]] : ''))];
    return `<tr>${cells.map((c) => `<td>${escapeHtml(c)}</td>`).join('')}</tr>`;
  });
  if (rows.length > shown.length) shown.push(`<tr><td colspan="${state.variables.length + 1}" class="empty-hint">... 외 ${rows.length - shown.length}행</td></tr>`);
  el.review.body.innerHTML = shown.join('');
  return confirmModal(el.review.modal, el.review.proceed, el.review.cancel);
}

/** 같은 이름의 파일이 이미 있으면 덮어쓸지 묻는다. 계속하면 true */
function confirmOverwrite(outputDir, filenames) {
  const existing = filenames.filter((name) => fs.existsSync(path.join(outputDir, name)));
  if (existing.length === 0) return true;
  const sample = existing.slice(0, 3).join(', ');
  return confirm(`선택한 폴더에 이름이 같은 파일이 ${existing.length}개 있습니다 (${sample}${existing.length > 3 ? ' 등' : ''}).\n덮어쓸까요?`);
}

el.btnGenerateBatch.addEventListener('click', async () => {
  try {
    if (!state.templatePath) throw new Error(NO_TEMPLATE);
    if (state.variables.length === 0) throw new Error(NO_VARIABLES);
    const { rows } = state.data;
    if (rows.length === 0) throw new Error('데이터 파일을 먼저 불러오세요.');
    const format = merge.templateFormat(state.templatePath);
    if (!state.variables.some((v) => state.mapping[v])) throw new Error('매핑된 변수가 하나도 없습니다. 컬럼 매핑을 확인하세요.');

    const filenames = merge.outputFilenames(rows, state.filenameColumn, format);
    if (!(await showReview(filenames))) {
      log('사용자가 확인 단계에서 취소했습니다. 매핑을 수정한 뒤 다시 시도하세요.', 'warn');
      return;
    }
    const outputDir = await ipcRenderer.invoke('dialog:selectFolder', { title: '결과 파일을 저장할 폴더 선택' });
    if (!outputDir) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }
    if (!confirmOverwrite(outputDir, filenames)) {
      log('기존 파일을 덮어쓰지 않도록 일괄 생성을 취소했습니다. 다른 폴더를 선택하세요.', 'warn');
      return;
    }

    log(`일괄 생성을 시작합니다... (${rows.length}건, 형식: ${format.toUpperCase()})`);
    const template = merge.openTemplate(state.templatePath);
    let succeeded = 0;
    const errors = [];
    for (let i = 0; i < rows.length; i++) {
      try {
        safeWriteFileSync(path.join(outputDir, filenames[i]), template.render(merge.rowData(rows[i], state.variables, state.mapping)));
        succeeded += 1;
      } catch (err) {
        errors.push({ row: i + 1, message: friendlyErrorMessage(err) });
      }
      const done = i + 1;
      if (done % PROGRESS_LOG_EVERY === 0 && done < rows.length) log(`진행 중... (${done}/${rows.length})`);
      if (done % YIELD_EVERY === 0) await new Promise((resolve) => setTimeout(resolve, 0));
    }

    log(`일괄 생성 완료: 총 ${rows.length}건 중 ${succeeded}건 성공${errors.length > 0 ? `, ${errors.length}건 실패` : ''}.`, errors.length > 0 ? 'warn' : 'ok');
    errors.slice(0, 5).forEach((e) => log(`  - ${e.row}행 실패: ${e.message}`, 'warn'));
    if (errors.length > 5) log(`  ... 외 ${errors.length - 5}건 실패`, 'warn');
    alert(
      errors.length === 0
        ? `✅ 총 ${succeeded}개의 문서를 생성했습니다.\n저장 위치: ${outputDir}`
        : `⚠️ ${succeeded}개 생성 완료, ${errors.length}건 실패했습니다.\n자세한 내용은 로그를 확인하세요.\n저장 위치: ${outputDir}`
    );
  } catch (err) {
    notifyError('일괄 생성 중 오류가 발생했습니다', err);
  }
});
