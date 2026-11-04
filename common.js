// common.js
// 화면 공통 도우미: 진행 로그, 오류 알림, 확인 창, 파일 끌어다 놓기, 시트(표) 고르기, 동의어 사전 창.
// 파일·데이터를 다루는 규칙은 화면 없는 모듈에 있다:
//   file-utils.js(파일·오류 문구) · data-file.js(데이터 파일 읽기) · synonyms.js(동의어 사전 저장) · auto-match.js(자동 매핑)

const { ipcRenderer, webUtils } = require('electron');
const path = require('path');
const { friendlyErrorMessage } = require('./file-utils.js');
const dataFile = require('./data-file.js');
const synonyms = require('./synonyms.js');
const autoMatch = require('./auto-match.js');
const { describeRecognition, formatPreviewValue } = require('./table-reader.js');

/* ---------- 진행 로그·오류 알림 ---------- */

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

document.getElementById('btnClearLog').addEventListener('click', () => {
  logArea.value = '';
});

function notifyError(context, err) {
  const msg = friendlyErrorMessage(err);
  log(`${context}: ${msg}`, 'error');
  alert(`❌ ${context}\n\n${msg}`);
}

/** 상태 배지(대기 중 · N개 발견 · 실패 …) */
function setPill(el, text, ok) {
  el.textContent = text;
  el.className = ok ? 'pill ok' : 'pill warn';
}

/* ---------- 확인 창 ---------- */

/**
 * "진행 / 취소" 확인 창을 띄우고 진행(true)·취소(false)를 Promise로 돌려준다. 바깥(배경)을 누르면 취소.
 * 창에 채울 내용(요약·미리보기 표)은 부르는 쪽이 만든다.
 */
function confirmModal(modalEl, proceedBtn, cancelBtn) {
  return new Promise((resolve) => {
    modalEl.hidden = false;
    const finish = (answer) => {
      modalEl.hidden = true;
      proceedBtn.onclick = null;
      cancelBtn.onclick = null;
      modalEl.onclick = null;
      resolve(answer);
    };
    proceedBtn.onclick = () => finish(true);
    cancelBtn.onclick = () => finish(false);
    modalEl.onclick = (e) => {
      if (e.target === modalEl) cancelBtn.click();
    };
  });
}

/* ---------- 파일 고르기 ---------- */

/** 드롭존(누르면 파일 열기 창, 끌어다 놓기도 됨). 고른 파일 경로를 onFileSelected(filePath)로 */
function bindDropzone(zoneEl, fileNameEl, filters, onFileSelected) {
  const select = (filePath) => {
    fileNameEl.textContent = path.basename(filePath);
    onFileSelected(filePath);
  };
  zoneEl.addEventListener('click', async () => {
    try {
      const paths = await ipcRenderer.invoke('dialog:openFile', { filters, multi: false });
      if (paths && paths.length > 0) select(paths[0]);
    } catch (err) {
      notifyError('파일 선택 중 오류가 발생했습니다', err);
    }
  });
  zoneEl.addEventListener('dragover', (e) => {
    e.preventDefault();
    zoneEl.classList.add('dragover');
  });
  zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('dragover'));
  zoneEl.addEventListener('drop', (e) => {
    e.preventDefault();
    zoneEl.classList.remove('dragover');
    try {
      const files = e.dataTransfer.files;
      if (!files || files.length === 0) return;
      // Electron 32부터 File.path가 없어져 webUtils로 실제 경로를 얻는다
      const filePath = webUtils.getPathForFile(files[0]);
      if (!filePath) throw new Error('드래그된 파일의 경로를 확인할 수 없습니다.');
      select(filePath);
    } catch (err) {
      notifyError('파일 드롭 처리 중 오류가 발생했습니다', err);
    }
  });
}

/* ---------- 데이터 파일 ---------- */

const noteToLog = (note) => log(note);

/** 데이터 파일의 시트(한글 문서는 표)를 읽는다. 한글 문서에서 찾은 표 목록은 로그에 */
const readSheetTable = (filePath, sheetName, options = {}) => dataFile.readSheetTable(filePath, sheetName, { ...options, onNote: noteToLog });

/**
 * 시트(표) 고르기 목록을 채운다. 처음 고른 시트는 onChange로, 사용자가 목록에서 바꾸면 onChange 다음에
 * onUserChangedSheet도 부른다('시트를 바꾸면 무엇을 비울지'는 부르는 쪽이 정한다).
 */
function refreshSheetSelector(filePath, sheetRowEl, sheetSelectEl, onChange, onUserChangedSheet) {
  try {
    const { names, unit, selected } = dataFile.sheetChoices(dataFile.readWorkbook(filePath, { onNote: noteToLog }));
    sheetSelectEl.innerHTML = '';
    names.forEach((name) => sheetSelectEl.appendChild(new Option(name, name)));
    // 한글 문서는 시트 대신 '표'를 고르므로 안내 문구도 그에 맞춘다
    const labelEl = sheetRowEl.querySelector('label');
    if (labelEl) labelEl.textContent = `${unit} 선택`;
    sheetRowEl.hidden = names.length <= 1;
    if (names.length > 1) log(`${unit} ${names.length}개 발견: ${names.join(', ')}`);

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

/* ---------- 동의어 사전 ---------- */

let userDataPath = null;
let synonymGroups = synonyms.DEFAULT_SYNONYM_GROUPS.map((g) => [...g]);

/** 템플릿 변수 ↔ 데이터 열 자동 매핑(사용자 동의어 사전 포함). 순서에 따라 결과가 달라지지 않게 한꺼번에 정한다 */
const autoMatchAll = (headersB, headersA) => autoMatch.autoMatchAll(headersB, headersA, synonymGroups);

const synonymModal = document.getElementById('synonymManagerModal');
const synonymTextarea = document.getElementById('synonymGroupsTextarea');
const synonymError = document.getElementById('synonymManagerError');
const closeSynonymModal = () => {
  synonymModal.hidden = true;
};

document.getElementById('btnManageSynonymsTab2').addEventListener('click', () => {
  synonymTextarea.value = synonyms.groupsToText(synonymGroups);
  synonymError.hidden = true;
  synonymModal.hidden = false;
  synonymTextarea.focus();
});
document.getElementById('synonymManagerCancel').addEventListener('click', closeSynonymModal);
synonymModal.addEventListener('click', (e) => {
  if (e.target === synonymModal) closeSynonymModal();
});
document.getElementById('synonymManagerReset').addEventListener('click', () => {
  if (!confirm('기본 동의어 사전으로 되돌릴까요? 직접 추가/수정한 내용은 사라집니다.')) return;
  synonymTextarea.value = synonyms.groupsToText(synonyms.DEFAULT_SYNONYM_GROUPS);
});
document.getElementById('synonymManagerSave').addEventListener('click', () => {
  try {
    synonymGroups = synonyms.textToGroups(synonymTextarea.value);
    synonyms.saveSynonymGroups(userDataPath, synonymGroups);
    closeSynonymModal();
    log(`동의어 사전을 저장했습니다 (${synonymGroups.length}개 그룹). 다음 [데이터 불러오기]부터 적용됩니다.`, 'ok');
  } catch (err) {
    synonymError.textContent = friendlyErrorMessage(err);
    synonymError.hidden = false;
  }
});

// 앱 시작 때 사용자 데이터 폴더를 받아 동의어 사전을 불러온다
ipcRenderer
  .invoke('app:getUserDataPath')
  .then((p) => {
    userDataPath = p;
    const loaded = synonyms.loadSynonymGroups(p);
    synonymGroups = loaded.groups;
    if (loaded.error) log(`동의어 사전을 불러오는 중 오류가 발생하여 기본값으로 되돌립니다: ${friendlyErrorMessage(loaded.error)}`, 'warn');
  })
  .catch((err) => {
    log(`앱 데이터 폴더 경로를 확인하지 못했습니다: ${friendlyErrorMessage(err)}`, 'warn');
  });

module.exports = {
  log,
  notifyError,
  setPill,
  confirmModal,
  bindDropzone,
  refreshSheetSelector,
  readSheetTable,
  describeRecognition,
  autoMatchAll,
  formatPreviewValue,
  DATA_FILE_FILTERS: dataFile.DATA_FILE_FILTERS,
};
