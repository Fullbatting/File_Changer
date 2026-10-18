// tab1.js
// TAB 1. 문서 양식 변환기: 구조 분석, 매핑 작업대(상태 모델·실행 취소·결과 미리보기·값 수정),
// 프리셋, 서식 적용 저장(ETL). 화면 그리기와 조작은 mapping-workbench.js가 맡고,
// 상태는 이 파일만 가진다(작업대는 api를 통해 읽고 바꾼다).

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const common = require('./common.js');
const { createWorkbench } = require('./mapping-workbench.js');

const {
  log,
  notifyError,
  friendlyErrorMessage,
  safeReadFileSync,
  safeWriteFileSync,
  confirmModal,
  showConfirmDialog,
  DATA_FILE_FILTERS,
  JSON_FILTERS,
  bindDropzone,
  refreshSheetSelector,
  extractHeaders,
  readSheetTable,
  findAutoMatch,
  detectValueKind,
  xmlEscape,
  formatPreviewValue,
  getUserDataPath,
  onUserDataPathReady,
} = common;

let fileAPath = null;
let fileBPath = null;
let sheetAName = null;
let sheetBName = null;

/* -------------------------------------------------------------------------
 * 상태 모델 (탭1의 유일한 진실 원천)
 * [구조 분석] 때 원본을 한 번 읽어 source에 둔다. 변환 때 다시 읽지 않으므로
 * 미리보기와 저장 결과가 어긋나지 않는다. 화면 요소에서 상태를 거꾸로 읽지 않고,
 * 모든 변경은 commit()을 거친다.
 * ------------------------------------------------------------------------- */

let source = null; // readSheetTable 결과 + { samples, valueKinds, emptyColumns, fileName, numericText }
let headersA = [];
let headersB = []; // 컬럼키 (같은 이름이 둘이면 "비고_1"처럼 구분됨)
let labelsB = []; // 결과 파일 헤더에 쓸 원래 컬럼 이름
let mapping = {}; // { B키: A키 | '' }
let matchReason = {}; // { B키: 'exact'|'synonym'|'fuzzy'|'confirmed'|'manual'|'preset'|'rule'|'none' }
let autoCandidates = {}; // { B키: [A키] } 자동 매핑이 모호해서 멈춘 후보
let requiredColumns = new Set(); // 저장 전 검증에서 "필수"로 지정된 B 컬럼
let overrides = new Map(); // "원본행인덱스|B키" -> 사용자가 미리보기에서 고친 글자
let history = { undo: [], redo: [] };
let derivedCache = null;
let presetSuggestionDismissed = false;

const HISTORY_LIMIT = 200;

const dropA = document.getElementById('dropA');
const dropB = document.getElementById('dropB');
const fileNameA = document.getElementById('fileNameA');
const fileNameB = document.getElementById('fileNameB');
const sheetRowA = document.getElementById('sheetRowA');
const sheetRowB = document.getElementById('sheetRowB');
const sheetSelectA = document.getElementById('sheetSelectA');
const sheetSelectB = document.getElementById('sheetSelectB');
const analyzeStatus = document.getElementById('analyzeStatus');
const presetSelect = document.getElementById('presetSelect');
const convertDesc = document.getElementById('convertDesc');

const overrideKey = (i, b) => `${i}|${b}`;
const overrideColumn = (key) => key.slice(key.indexOf('|') + 1);
const isEmptyValue = (v) => v === '' || v === null || v === undefined;

function snapshot() {
  return {
    mapping: { ...mapping },
    matchReason: { ...matchReason },
    required: Array.from(requiredColumns),
    overrides: Array.from(overrides),
  };
}

function restore(state) {
  mapping = { ...state.mapping };
  matchReason = { ...state.matchReason };
  requiredColumns = new Set(state.required);
  overrides = new Map(state.overrides);
  invalidate();
}

function sameState(a, b) {
  return (
    JSON.stringify(a.mapping) === JSON.stringify(b.mapping) &&
    JSON.stringify(a.matchReason) === JSON.stringify(b.matchReason) &&
    JSON.stringify([...a.required].sort()) === JSON.stringify([...b.required].sort()) &&
    JSON.stringify(a.overrides) === JSON.stringify(b.overrides)
  );
}

function invalidate() {
  derivedCache = null;
}

/**
 * 모든 상태 변경의 단일 진입점. mutate(draft)로 바뀐 상태를 만든 뒤,
 * 연결이 바뀌는 열에 직접 고친 값이 있으면 "지워진다"는 확인을 받고 그 열의 고친 값만
 * 지운다. 그다음 실행 취소 기록을 남기고 작업대를 다시 그린다. 바뀐 게 없거나
 * 확인에서 취소하면 false.
 */
async function commit(label, mutate) {
  if (!source) return false;
  const before = snapshot();
  const draft = {
    mapping: { ...mapping },
    matchReason: { ...matchReason },
    required: new Set(requiredColumns),
    overrides: new Map(overrides),
  };
  mutate(draft);

  const changedColumns = new Set(headersB.filter((b) => (draft.mapping[b] || '') !== (before.mapping[b] || '')));
  if (changedColumns.size > 0) {
    const affected = Array.from(draft.overrides.keys()).filter((k) => changedColumns.has(overrideColumn(k)));
    if (affected.length > 0) {
      const columnLabels = Array.from(new Set(affected.map(overrideColumn))).map(labelOfB);
      const ok = await showConfirmDialog({
        title: '연결을 바꿀까요?',
        message: `${columnLabels.join(', ')} 열에서 직접 고친 값 ${affected.length}개가 지워집니다. 계속할까요?`,
        note: '바꾼 뒤에도 실행 취소로 되돌릴 수 있습니다.',
        okText: '지우고 바꾸기',
        danger: true,
      });
      if (!ok) return false;
      affected.forEach((k) => draft.overrides.delete(k));
    }
  }

  const after = {
    mapping: draft.mapping,
    matchReason: draft.matchReason,
    required: Array.from(draft.required),
    overrides: Array.from(draft.overrides),
  };
  if (sameState(before, after)) return false;

  history.undo.push({ label, state: before });
  if (history.undo.length > HISTORY_LIMIT) history.undo.shift();
  history.redo = [];
  restore(after);
  workbench.render();
  return true;
}

function undo() {
  const entry = history.undo.pop();
  if (!entry) return false;
  history.redo.push({ label: entry.label, state: snapshot() });
  restore(entry.state);
  // 실행 취소·다시 실행은 로그에 남기지 않는다(로그창 갱신이 조작마다 화면을 무겁게 하고, 화면에 바로 보이므로)
  workbench.render();
  return true;
}

function redo() {
  const entry = history.redo.pop();
  if (!entry) return false;
  history.undo.push({ label: entry.label, state: snapshot() });
  restore(entry.state);
  workbench.render();
  return true;
}

function labelOfB(b) {
  const i = headersB.indexOf(b);
  return i >= 0 ? labelsB[i] : b;
}

/* ---------- 파생 데이터: 미리보기 · 검증 · 요약 · 저장이 모두 이것 하나를 쓴다 ---------- */

/** 값이 "숫자로 바꿔도 안전한" 문자열인지 (isSafeNumericString 설명 참고) */
function isSafeNumericString(v) {
  if (typeof v === 'number') return Number.isFinite(v);
  if (typeof v !== 'string') return false;
  const trimmed = v.trim();
  if (trimmed === '' || !/^-?\d+(\.\d+)?$/.test(trimmed)) return false;
  return String(Number(trimmed)) === trimmed;
}

/** 사용자가 고친 글자를 숫자로 볼 수 있으면 그 숫자를, 아니면 null. 천 단위 쉼표("1,500,000")도 받는다 */
function parseOverrideNumber(str) {
  const s = String(str).trim();
  const plain = /^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s) ? s.replace(/,/g, '') : s;
  return isSafeNumericString(plain) ? Number(plain) : null;
}

/** "YYYY-MM-DD" 또는 "YYYY-MM-DD HH:MM(:SS)"만 현지 시각 Date로. 그 외에는 null */
function parseOverrideDate(str) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(String(str).trim());
  if (!m) return null;
  const [y, mo, d, h = 0, mi = 0, s = 0] = m.slice(1).map((x) => (x === undefined ? 0 : Number(x)));
  const date = new Date(y, mo - 1, d, h, mi, s);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d || h > 23 || mi > 59 || s > 59) return null;
  return date;
}

/** 고친 값(글자)을 열 종류에 맞춰 저장할 값으로. 숫자·날짜로 볼 수 없으면 그 칸만 글자로 둔다 */
function coerceOverride(str, kind) {
  if (str === '') return '';
  if (kind === 'number') {
    const n = parseOverrideNumber(str);
    return n === null ? str : n;
  }
  if (kind === 'date') {
    const d = parseOverrideDate(str);
    return d || str;
  }
  return str;
}

/** 저장할 때 이 칸을 숫자로 쓸지 (숫자 열 안에서도 글자로 고친 칸은 글자로 둔다) */
function isNumericCell(v, numericText) {
  return typeof v === 'number' ? Number.isFinite(v) : numericText && isSafeNumericString(v);
}

/** 중복 행 비교용 값 표시(이전 검증 규칙과 같이 화면 글자 기준으로 비교한다) */
function valueToken(v) {
  if (v instanceof Date) return formatPreviewValue(v);
  return isEmptyValue(v) ? '' : String(v);
}

/*
 * 원본 컬럼별 통계 캐시. 컬럼 하나를 한 번만 훑어 값 목록·값 번호(중복 비교용)·빈칸·숫자/날짜
 * 여부를 기억해 두고, 연결을 바꾸거나 값을 고칠 때는 바뀐 부분만 다시 계산한다(1만 행에서도
 * 조작마다 전체 표를 다시 만들지 않기 위함). [구조 분석] 때마다 비운다.
 */
let columnStatsCache = new Map();
const EMPTY_COLUMN_KEY = '\u0000빈 열';

function columnStats(a) {
  const key = a || EMPTY_COLUMN_KEY;
  if (columnStatsCache.has(key)) return columnStatsCache.get(key);
  const n = source.rows.length;
  const values = new Array(n);
  const ids = new Int32Array(n);
  const idMap = new Map([['', 0]]);
  const empty = new Uint8Array(n);
  const numeric = new Uint8Array(n); // 1: 숫자, 2: 소수가 있는 숫자
  const date = new Uint8Array(n);
  let nonEmptyCount = 0;
  let numericCount = 0;
  let dateCount = 0;
  let decimalCount = 0;
  for (let i = 0; i < n; i++) {
    const raw = a ? source.rows[i][a] : '';
    const v = raw === undefined ? '' : raw;
    values[i] = v;
    if (isEmptyValue(v)) {
      empty[i] = 1;
      continue;
    }
    nonEmptyCount += 1;
    const token = valueToken(v);
    let id = idMap.get(token);
    if (id === undefined) {
      id = idMap.size;
      idMap.set(token, id);
    }
    ids[i] = id;
    if (isNumericCell(v, source.numericText)) {
      const decimal = !Number.isInteger(Number(v));
      numeric[i] = decimal ? 2 : 1;
      numericCount += 1;
      if (decimal) decimalCount += 1;
    }
    if (v instanceof Date) {
      date[i] = 1;
      dateCount += 1;
    }
  }
  const stats = { values, ids, idMap, empty, numeric, date, nonEmptyCount, numericCount, dateCount, decimalCount };
  columnStatsCache.set(key, stats);
  return stats;
}

/**
 * 열 종류 판정(저장과 미리보기가 같이 쓴다). 고친 칸은 빼고 원본에서 온 값만으로 판정하므로
 * 고친 값 하나 때문에 열 전체의 숫자·날짜 판정이 바뀌지 않는다.
 * - number: 빈 값을 뺀 모든 값이 숫자(엑셀 원본은 원본에서 숫자였던 값만, numericText 원본은
 *   "안전한 숫자 글자"도) — buildStyledExcelBuffer의 숫자 열 규칙과 같다.
 * - date: 빈 값을 뺀 모든 값이 날짜.
 * 원본에서 온 값이 하나도 없는 열(연결 없음 등)은 고친 값만으로 판정한다.
 */
function columnKind(stats, overriddenRows, overrideTexts) {
  let nonEmpty = stats.nonEmptyCount;
  let numeric = stats.numericCount;
  let date = stats.dateCount;
  if (overriddenRows) {
    overriddenRows.forEach((i) => {
      if (!stats.empty[i]) nonEmpty -= 1;
      if (stats.numeric[i]) numeric -= 1;
      if (stats.date[i]) date -= 1;
    });
  }
  if (nonEmpty > 0) return numeric === nonEmpty ? 'number' : date === nonEmpty ? 'date' : 'text';
  const texts = (overrideTexts || []).filter((v) => !isEmptyValue(v));
  if (texts.length > 0 && texts.every((v) => parseOverrideNumber(v) !== null)) return 'number';
  if (texts.length > 0 && texts.every((v) => parseOverrideDate(v) !== null)) return 'date';
  return 'text';
}

/**
 * 파생 데이터를 계산한다(상태가 바뀔 때만 다시 계산).
 * value(i, b) = B 컬럼 칸 값: 고친 값이 있으면 열 종류에 맞춰 바꾼 그 값, 아니면 원본 값,
 * 연결이 없으면 ''. 저장 파일은 getRows()(= 모든 행의 value) 그대로다.
 */
function getDerived() {
  if (derivedCache) return derivedCache;
  if (!source) return null;
  const n = source.rows.length;

  // 고친 값을 열별로 모은다
  const overridesByColumn = new Map(); // b -> Map(i -> 글자)
  overrides.forEach((text, key) => {
    const i = Number(key.slice(0, key.indexOf('|')));
    const b = overrideColumn(key);
    if (!overridesByColumn.has(b)) overridesByColumn.set(b, new Map());
    overridesByColumn.get(b).set(i, text);
  });

  const statsList = headersB.map((b) => columnStats(mapping[b]));
  const kinds = [];
  const typedOverrides = headersB.map(() => null); // ci -> Map(i -> 바꾼 값)
  headersB.forEach((b, ci) => {
    const stats = statsList[ci];
    const colOverrides = overridesByColumn.get(b);
    const kind = columnKind(stats, colOverrides && Array.from(colOverrides.keys()), colOverrides && Array.from(colOverrides.values()));
    let decimal = stats.decimalCount;
    if (colOverrides) {
      const typed = new Map();
      colOverrides.forEach((text, i) => {
        const v = coerceOverride(text, kind);
        typed.set(i, v);
        if (stats.numeric[i] === 2) decimal -= 1;
        if (isNumericCell(v, source.numericText) && !Number.isInteger(Number(v))) decimal += 1;
      });
      typedOverrides[ci] = typed;
    }
    kinds.push({ kind, hasDecimal: kind === 'number' && decimal > 0 });
  });

  const columnIndex = new Map(headersB.map((b, ci) => [b, ci]));
  const value = (i, b) => {
    const ci = columnIndex.get(b);
    if (ci === undefined) return '';
    const typed = typedOverrides[ci];
    if (typed && typed.has(i)) return typed.get(i);
    return statsList[ci].values[i];
  };

  // 검증: 필수 빈칸
  const requiredIdx = headersB.map((b, ci) => (requiredColumns.has(b) ? ci : -1)).filter((ci) => ci >= 0);
  const missingByColumn = {};
  const missingByRow = new Array(n).fill(null);
  let missingRowCount = 0;
  requiredIdx.forEach((ci) => {
    const b = headersB[ci];
    const stats = statsList[ci];
    const typed = typedOverrides[ci];
    let count = 0;
    for (let i = 0; i < n; i++) {
      const isEmpty = typed && typed.has(i) ? isEmptyValue(typed.get(i)) : stats.empty[i] === 1;
      if (!isEmpty) continue;
      count += 1;
      if (missingByRow[i]) missingByRow[i].push(b);
      else {
        missingByRow[i] = [b];
        missingRowCount += 1;
      }
    }
    missingByColumn[b] = count;
  });

  // 검증: 모든 값이 같은 중복 행. 칸마다 값 번호를 매겨 두었으므로 행을 숫자 묶음으로 비교한다
  const overrideIds = headersB.map((b, ci) => {
    const typed = typedOverrides[ci];
    if (!typed) return null;
    const ids = new Map();
    typed.forEach((v, i) => {
      const token = valueToken(v);
      const known = statsList[ci].idMap.get(token);
      ids.set(i, known !== undefined ? known : `새:${token}`);
    });
    return ids;
  });
  const idAt = (ci, i) => {
    const ov = overrideIds[ci];
    if (ov && ov.has(i)) return ov.get(i);
    return statsList[ci].ids[i];
  };
  const sameRow = (x, y) => {
    for (let ci = 0; ci < headersB.length; ci++) if (idAt(ci, x) !== idAt(ci, y)) return false;
    return true;
  };
  const duplicateOf = new Int32Array(n).fill(-1);
  const buckets = new Map(); // 해시 -> 처음 나온 행들
  let duplicateCount = 0;
  const duplicateExamples = [];
  for (let i = 0; i < n; i++) {
    let hash = 17;
    for (let ci = 0; ci < headersB.length; ci++) {
      const id = idAt(ci, i);
      const num = typeof id === 'number' ? id : 1000003 + id.length * 31 + id.charCodeAt(id.length - 1);
      hash = (Math.imul(hash, 31) + num) | 0;
    }
    const bucket = buckets.get(hash);
    if (!bucket) {
      buckets.set(hash, [i]);
      continue;
    }
    const first = bucket.find((j) => sameRow(j, i));
    if (first === undefined) {
      bucket.push(i);
    } else {
      duplicateOf[i] = first;
      duplicateCount += 1;
      if (duplicateExamples.length < 5) duplicateExamples.push({ row: i, duplicateOf: first });
    }
  }

  const overrideCountByColumn = {};
  overridesByColumn.forEach((m, b) => {
    overrideCountByColumn[b] = m.size;
  });

  // 열 상태(열 지도·문제 있는 열만): 필수 빈칸 > 확인 필요 > 연결 없음 > 고친 값 > 정상
  const columnStatus = {};
  headersB.forEach((b) => {
    if (requiredColumns.has(b) && missingByColumn[b] > 0) columnStatus[b] = 'reqmissing';
    else if (matchReason[b] === 'fuzzy') columnStatus[b] = 'fuzzy';
    else if (!mapping[b]) columnStatus[b] = 'none';
    else if (overrideCountByColumn[b]) columnStatus[b] = 'edited';
    else columnStatus[b] = 'ok';
  });

  let rowsCache = null;
  derivedCache = {
    rowCount: n,
    value,
    /** 저장용 전체 행 { B키: 값 } — 필요할 때 한 번만 만든다 */
    getRows() {
      if (!rowsCache) {
        rowsCache = new Array(n);
        for (let i = 0; i < n; i++) {
          const obj = {};
          headersB.forEach((b) => {
            obj[b] = value(i, b);
          });
          rowsCache[i] = obj;
        }
      }
      return rowsCache;
    },
    kinds,
    missingByColumn,
    missingByRow,
    missingRowCount,
    duplicateOf,
    duplicateCount,
    duplicateExamples,
    overrideCountByColumn,
    columnStatus,
  };
  return derivedCache;
}

/** 저장될 엑셀 화면과 같은 글자로 보여준다(날짜 yyyy-mm-dd, 숫자 열은 원본 숫자 형식 또는 천 단위 구분) */
function displayValue(v, b) {
  if (isEmptyValue(v)) return '';
  if (v instanceof Date) return formatPreviewValue(v);
  const derived = getDerived();
  const ci = headersB.indexOf(b);
  const kind = derived && ci >= 0 ? derived.kinds[ci] : null;
  if (kind && kind.kind === 'number' && isNumericCell(v, source.numericText)) {
    const a = mapping[b];
    const format = (a && source.numberFormats[a]) || (kind.hasDecimal ? '#,##0.######' : '#,##0');
    try {
      return XLSX.SSF.format(format, Number(v));
    } catch (e) {
      return String(v);
    }
  }
  return String(v);
}

/** 미리보기 칸 정렬: 날짜는 가운데('date'), 숫자로 저장될 칸은 오른쪽('num'), 나머지는 '' */
function cellAlign(v, b) {
  if (v instanceof Date) return 'date';
  if (isEmptyValue(v)) return '';
  const derived = getDerived();
  const ci = headersB.indexOf(b);
  return derived && ci >= 0 && derived.kinds[ci].kind === 'number' && isNumericCell(v, source.numericText) ? 'num' : '';
}

/** 고치기 전 원래 값(원본에서 온 값) */
function originalValue(i, b) {
  const a = mapping[b];
  if (!a) return '';
  const v = source.rows[i][a];
  return v === undefined ? '' : v;
}

/* ---------- 작업대에서 쓰는 조작 (모두 commit을 거친다) ---------- */

const api = {
  getState() {
    return {
      source,
      headersA,
      headersB,
      labelsB,
      mapping,
      matchReason,
      autoCandidates,
      required: requiredColumns,
      overrides,
      canUndo: history.undo.length > 0,
      canRedo: history.redo.length > 0,
      sourceName: fileAPath ? path.basename(fileAPath) : '',
      targetName: fileBPath ? path.basename(fileBPath) : '',
    };
  },
  getDerived,
  displayValue,
  cellAlign,
  originalValue,
  labelOfB,
  connect(b, a) {
    return commit(`${labelOfB(b)} ← ${a} 연결`, (d) => {
      d.mapping[b] = a;
      d.matchReason[b] = 'manual';
    });
  },
  disconnect(b) {
    return commit(`${labelOfB(b)} 연결 끊기`, (d) => {
      d.mapping[b] = '';
      d.matchReason[b] = 'none';
    });
  },
  confirmMatch(b) {
    return commit(`${labelOfB(b)} 연결 확인`, (d) => {
      if (d.matchReason[b] === 'fuzzy') d.matchReason[b] = 'confirmed';
    });
  },
  confirmAllFuzzy(list) {
    return commit(`비슷한 이름 연결 ${list.length}개 확인`, (d) => {
      list.forEach((b) => {
        if (d.matchReason[b] === 'fuzzy') d.matchReason[b] = 'confirmed';
      });
    });
  },
  toggleRequired(b) {
    const on = !requiredColumns.has(b);
    return commit(`${labelOfB(b)} 필수 ${on ? '지정' : '해제'}`, (d) => {
      if (on) d.required.add(b);
      else d.required.delete(b);
    });
  },
  setRequired(b, on) {
    return commit(`${labelOfB(b)} 필수 ${on ? '지정' : '해제'}`, (d) => {
      if (on) d.required.add(b);
      else d.required.delete(b);
    });
  },
  /** 값 고치기. 원래 값과 같은 글자로 고치면 고친 값을 지운다(원래 값으로 돌아감) */
  setOverride(i, b, text) {
    const original = formatPreviewValue(originalValue(i, b));
    return commit(`${source.rowNumbers[i]}행 ${labelOfB(b)} 고치기`, (d) => {
      if (text === original) d.overrides.delete(overrideKey(i, b));
      else d.overrides.set(overrideKey(i, b), text);
    });
  },
  clearOverride(i, b) {
    return commit(`${source.rowNumbers[i]}행 ${labelOfB(b)} 원래 값으로`, (d) => {
      d.overrides.delete(overrideKey(i, b));
    });
  },
  async clearAllOverrides() {
    if (overrides.size === 0) return false;
    const ok = await showConfirmDialog({
      title: '고친 값을 모두 되돌릴까요?',
      message: `직접 고친 값 ${overrides.size}개를 모두 원래 값으로 되돌립니다.`,
      note: '실행 취소로 다시 살릴 수 있습니다.',
      okText: '모두 되돌리기',
      danger: true,
    });
    if (!ok) return false;
    return commit('고친 값 모두 되돌리기', (d) => d.overrides.clear());
  },
  getOverride(i, b) {
    return overrides.get(overrideKey(i, b));
  },
  undo,
  redo,
  getPresetSuggestion,
  dismissPresetSuggestion() {
    presetSuggestionDismissed = true;
    workbench.render();
  },
  applyPreset(id) {
    return applyPresetById(id);
  },
};

const workbench = createWorkbench(api);

/* ---------- 파일·시트 선택 ---------- */

/** 고친 값이 있으면 "사라진다"는 확인을 받는다. 고친 값이 없으면 바로 true */
function confirmDiscardOverrides(title) {
  if (overrides.size === 0) return Promise.resolve(true);
  return showConfirmDialog({
    title,
    message: `직접 고친 값 ${overrides.size}개가 모두 사라집니다. 계속할까요?`,
    note: '이 작업은 실행 취소로 되돌릴 수 없습니다.',
    okText: '계속',
    danger: true,
  });
}

// 시트 선택 <select> 클릭이 드롭존의 파일선택 클릭으로 전파되지 않도록 차단
[sheetRowA, sheetRowB].forEach((row) => {
  row.addEventListener('click', (e) => e.stopPropagation());
});

/** 파일이 (재)선택되면 이전 분석 결과가 새 파일과 어긋나지 않도록 상태 전체를 초기화 */
function resetAnalysisState() {
  source = null;
  headersA = [];
  headersB = [];
  labelsB = [];
  mapping = {};
  matchReason = {};
  autoCandidates = {};
  requiredColumns = new Set();
  overrides = new Map();
  history = { undo: [], redo: [] };
  presetSuggestionDismissed = false;
  currentPresetId = null;
  invalidate();
  columnStatsCache = new Map();
  workbench.reset();
  workbench.render();
  if (presetSelect) presetSelect.value = '';
  analyzeStatus.textContent = '대기 중';
  analyzeStatus.className = 'pill warn';
}

/** 탭1 A/B 시트를 사용자가 직접 바꿨을 때: 고친 값이 있으면 확인 후, 분석 결과를 초기화 */
async function onTab1SheetChanged(which, name) {
  const select = which === 'A' ? sheetSelectA : sheetSelectB;
  const ok = await confirmDiscardOverrides('시트를 바꿀까요?');
  if (!ok) {
    select.value = select.dataset.committed || '';
    if (which === 'A') sheetAName = select.value;
    else sheetBName = select.value;
    return;
  }
  select.dataset.committed = name;
  resetAnalysisState();
  log(`시트가 "${name}"(으)로 변경되었습니다. [구조 분석]을 다시 실행하세요.`, 'warn');
}

function bindTab1File(which, dropEl, fileNameEl, sheetRowEl, sheetSelectEl) {
  bindDropzone(dropEl, fileNameEl, DATA_FILE_FILTERS, async (p) => {
    const previous = which === 'A' ? fileAPath : fileBPath;
    const ok = await confirmDiscardOverrides(which === 'A' ? '원본 파일을 바꿀까요?' : '타겟 양식 파일을 바꿀까요?');
    if (!ok) {
      fileNameEl.textContent = previous ? path.basename(previous) : '';
      return;
    }
    if (which === 'A') fileAPath = p;
    else fileBPath = p;
    log(`${which === 'A' ? '원본(A)' : '타겟(B)'} 파일 선택됨: ${p}`);
    resetAnalysisState();
    refreshSheetSelector(
      p,
      sheetRowEl,
      sheetSelectEl,
      (name) => {
        if (which === 'A') sheetAName = name;
        else sheetBName = name;
      },
      (name) => onTab1SheetChanged(which, name)
    );
    sheetSelectEl.dataset.committed = sheetSelectEl.value;
  });
}

bindTab1File('A', dropA, fileNameA, sheetRowA, sheetSelectA);
bindTab1File('B', dropB, fileNameB, sheetRowB, sheetSelectB);

/** 원본 표에 화면용 정보(컬럼별 앞쪽 값, 값 모양, 값이 모두 빈 컬럼)를 붙인다 */
function prepareSource(table, filePath) {
  const samples = {};
  const valueKinds = {};
  const emptyColumns = new Set();
  table.headers.forEach((a) => {
    const values = [];
    for (let i = 0; i < table.rows.length && values.length < 20; i++) {
      const v = table.rows[i][a];
      if (!isEmptyValue(v)) values.push(v);
    }
    samples[a] = values;
    valueKinds[a] = detectValueKind(values);
    if (values.length === 0) emptyColumns.add(a);
  });
  return {
    ...table,
    samples,
    valueKinds,
    emptyColumns,
    // CSV·한글 문서 표는 모든 값이 글자로 들어오므로 숫자 글자도 숫자로 본다
    numericText: ['.csv', '.hwp', '.hwpx', '.hwpt'].includes(path.extname(filePath).toLowerCase()),
  };
}

/** 분석 직후 자동 매핑. 동의어/비슷한 이름 연결은 로그로도 알린다 */
function runAutoMapping() {
  mapping = {};
  matchReason = {};
  autoCandidates = {};
  const usedHeadersA = new Set();
  headersB.forEach((bCol) => {
    const { match, reason, candidates } = findAutoMatch(bCol, headersA, usedHeadersA);
    if (match) {
      mapping[bCol] = match;
      matchReason[bCol] = reason;
      usedHeadersA.add(match);
      if (reason === 'synonym') {
        log(`"${bCol}" 컬럼을 동의어 사전에 따라 "${match}" 컬럼에 자동 매핑했습니다.`);
      } else if (reason === 'fuzzy') {
        log(`"${bCol}" 컬럼을 이름이 비슷한 "${match}" 컬럼에 자동 매핑했습니다. 맞는지 확인하세요.`, 'warn');
      }
    } else {
      mapping[bCol] = '';
      matchReason[bCol] = 'none';
      if (candidates && candidates.length > 1) autoCandidates[bCol] = candidates;
    }
  });
}

document.getElementById('btnAnalyze').addEventListener('click', async () => {
  try {
    if (!fileAPath || !fileBPath) {
      throw new Error('원본(A) 파일과 타겟 양식(B) 파일을 모두 선택하세요.');
    }
    if (source && !(await confirmDiscardOverrides('다시 분석할까요?'))) return;
    log(`구조 분석을 시작합니다... (A 시트: ${sheetAName || '첫 번째 시트'}, B 시트: ${sheetBName || '첫 번째 시트'})`);
    const resultA = readSheetTable(fileAPath, sheetAName);
    const resultB = extractHeaders(fileBPath, sheetBName);
    [
      ['원본(A)', resultA.headerRowIndex],
      ['타겟(B)', resultB.headerRowIndex],
    ].forEach(([label, idx]) => {
      if (idx > 0) log(`${label} 파일: 제목/빈 행으로 보이는 ${idx}개 행을 건너뛰고 ${idx + 1}번째 행을 헤더로 인식했습니다.`, 'warn');
    });
    [
      ['원본(A)', resultA],
      ['타겟(B)', resultB],
    ].forEach(([label, { headers, labels }]) => {
      const renamed = headers.filter((k, i) => k !== labels[i]);
      if (renamed.length > 0) log(`${label} 파일에 이름이 같은 컬럼이 있어 ${renamed.join(', ')}(으)로 구분했습니다.`, 'warn');
    });

    resetAnalysisState();
    source = prepareSource(resultA, fileAPath);
    headersA = resultA.headers;
    headersB = resultB.headers;
    labelsB = resultB.labels;
    log(`원본(A) 컬럼 ${headersA.length}개, 타겟(B) 컬럼 ${headersB.length}개 발견`, 'ok');

    runAutoMapping();
    invalidate();
    workbench.onAnalyzed();
    workbench.render();
    analyzeStatus.textContent = '분석 완료';
    analyzeStatus.className = 'pill ok';
    log('매핑 작업대가 준비되었습니다. 연결과 결과 미리보기를 확인하세요.', 'ok');
    const suggestion = getPresetSuggestion();
    if (suggestion) log(`저장된 프리셋 "${suggestion.preset.name}"이(가) 이 양식과 ${suggestion.total}열 중 ${suggestion.matched}열 같습니다.`);
  } catch (err) {
    analyzeStatus.textContent = '분석 실패';
    analyzeStatus.className = 'pill warn';
    notifyError('구조 분석 중 오류가 발생했습니다', err);
  }
});

/* ---------- 매핑 규칙 파일 · 프리셋 ---------- */

document.getElementById('btnSaveMapping').addEventListener('click', async () => {
  try {
    if (headersB.length === 0) throw new Error('저장할 매핑 규칙이 없습니다. 먼저 구조 분석을 실행하세요.');

    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '매핑 규칙 저장',
      defaultPath: 'mapping_rule.json',
      filters: JSON_FILTERS,
    });
    if (!savePath) return;

    // 고친 값은 그 파일 하나에만 해당하므로 저장하지 않는다(연결과 필수 여부만)
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
 * 불러온 매핑({ B: A })과 필수 컬럼을 하나의 실행 취소 단계로 적용한다. 현재 원본(A)에
 * 없는 컬럼 값은 비우고 경고 로그를 남긴다. 매핑 규칙 파일과 프리셋이 함께 쓴다.
 */
async function applyLoadedMapping(label, reason, loadedMapping, loadedRequired) {
  let appliedCount = 0;
  let skippedCount = 0;
  const skipped = [];
  const ok = await commit(label, (d) => {
    Object.keys(loadedMapping).forEach((bCol) => {
      if (!headersB.includes(bCol)) return;
      const aCol = loadedMapping[bCol] || '';
      const exists = !aCol || headersA.includes(aCol);
      d.mapping[bCol] = exists ? aCol : '';
      d.matchReason[bCol] = d.mapping[bCol] ? reason : 'none';
      appliedCount += 1;
      if (aCol && !exists) {
        skippedCount += 1;
        skipped.push([bCol, aCol]);
      }
    });
    if (Array.isArray(loadedRequired)) d.required = new Set(loadedRequired.filter((c) => headersB.includes(c)));
  });
  skipped.forEach(([bCol, aCol]) => log(`"${bCol}" 컬럼의 매핑 값("${aCol}")이 현재 원본(A) 파일에 없어 초기화했습니다.`, 'warn'));
  return { ok, appliedCount, skippedCount };
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

    const { ok, appliedCount, skippedCount } = await applyLoadedMapping('매핑 규칙 불러오기', 'rule', loadedMapping, payload.requiredColumns);
    if (!ok) return;
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
const PRESET_SUGGEST_RATIO = 0.8;

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

/**
 * 분석한 타겟 양식과 B 컬럼 이름이 가장 많이 겹치는 저장된 프리셋(80% 이상)을 찾는다.
 * 작업대 위에 띠로 알릴 뿐 자동으로 적용하지는 않는다. 이미 적용 중이거나 닫았으면 null.
 */
function getPresetSuggestion() {
  if (!source || presetSuggestionDismissed || currentPresetId || headersB.length === 0) return null;
  let best = null;
  presets.forEach((preset) => {
    const presetTargets = new Set(Array.isArray(preset.targetHeaders) ? preset.targetHeaders : Object.keys(preset.mapping || {}));
    const matched = headersB.filter((b) => presetTargets.has(b)).length;
    if (matched / headersB.length >= PRESET_SUGGEST_RATIO && (!best || matched > best.matched)) {
      best = { preset, matched, total: headersB.length };
    }
  });
  return best;
}

async function applyPresetById(id) {
  const preset = presets.find((p) => p.id === id);
  if (!preset) return false;
  const { ok, appliedCount, skippedCount } = await applyLoadedMapping(`프리셋 "${preset.name}" 적용`, 'preset', preset.mapping || {}, preset.requiredColumns);
  if (!ok) return false;
  currentPresetId = id;
  presetSuggestionDismissed = true;
  renderPresetSelect();
  workbench.render();
  log(`프리셋 "${preset.name}"을(를) 적용했습니다. (적용 ${appliedCount}건, 불일치 ${skippedCount}건)`, 'ok');
  return true;
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
  presetSelect.addEventListener('change', async (e) => {
    const id = e.target.value;
    if (!id) {
      currentPresetId = null;
      return;
    }
    if (headersB.length === 0) {
      alert('먼저 원본(A)/타겟(B) 파일을 선택하고 [구조 분석]을 실행하세요.');
      e.target.value = currentPresetId || '';
      return;
    }
    const applied = await applyPresetById(id);
    if (!applied) e.target.value = currentPresetId || '';
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

    // 이름을 바꿔 저장하면 새 프리셋으로 만든다 — 선택돼 있던 프리셋을 새 이름으로 덮어쓰면 원래 프리셋이 사라진다
    const targetId = existing ? existing.id : generatePresetId();
    const now = new Date().toISOString();
    const existingRecord = presets.find((p) => p.id === targetId);

    // 고친 값은 저장하지 않는다(연결과 필수 여부만)
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
    workbench.render();
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

/* ---------- 변환 실행 및 저장 ---------- */

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

// 변환 결과 확인(요약) 모달
const convertReviewModal = document.getElementById('convertReviewModal');
const convertReviewSummary = document.getElementById('convertReviewSummary');
const convertReviewPreview = document.getElementById('convertReviewPreview');
const convertReviewTableHead = document.getElementById('convertReviewTableHead');
const convertReviewTableBody = document.getElementById('convertReviewTableBody');
const CONVERT_PREVIEW_ROW_LIMIT = 5;
const convertReviewProceedBtn = document.getElementById('convertReviewProceed');
const convertReviewCancelBtn = document.getElementById('convertReviewCancel');

/**
 * 변환 확인 모달: 요약을 보여준다. 시각화가 켜져 있으면 미리보기는 이미 화면에 있으므로 요약만,
 * 꺼져 있으면 저장될 표의 앞 5행도 함께 보여준다.
 * 문제가 있어도 저장을 막지는 않고 사용자가 확인 후 진행/취소를 선택한다
 * (CSV 임포트 도구들의 allowInvalidSubmit 기본 동작과 같은 취지).
 */
function showConvertReviewModal(derived) {
  const lines = [];
  lines.push(`<div class="review-line ok">총 <b>${derived.rowCount}</b>행을 저장합니다.</div>`);

  const missingEntries = Object.entries(derived.missingByColumn).filter(([, count]) => count > 0);
  if (missingEntries.length > 0) {
    missingEntries.forEach(([col, count]) => {
      lines.push(`<div class="review-line warn">필수 칸 "${xmlEscape(labelOfB(col))}"이(가) 비어 있는 행 ${count}행</div>`);
    });
  } else if (Object.keys(derived.missingByColumn).length > 0) {
    lines.push(`<div class="review-line ok">필수 칸에 빈 값이 없습니다.</div>`);
  }

  if (derived.duplicateCount > 0) {
    const examples = derived.duplicateExamples
      .map((e) => `${source.rowNumbers[e.row]}행(${source.rowNumbers[e.duplicateOf]}행과 같음)`)
      .join(', ');
    lines.push(`<div class="review-line warn">모든 값이 같은 중복 행 ${derived.duplicateCount}행: ${xmlEscape(examples)}${derived.duplicateCount > derived.duplicateExamples.length ? ' 등' : ''}</div>`);
  } else {
    lines.push(`<div class="review-line ok">완전히 중복된 행이 없습니다.</div>`);
  }

  const fuzzy = headersB.filter((b) => matchReason[b] === 'fuzzy');
  if (fuzzy.length > 0) {
    const detail = fuzzy.map((b) => `${labelOfB(b)} ← ${mapping[b]}`).join(', ');
    lines.push(`<div class="review-line warn">확인 안 한 비슷한 이름 연결 ${fuzzy.length}개 (${xmlEscape(detail)})</div>`);
  }
  if (overrides.size > 0) {
    lines.push(`<div class="review-line info">직접 고친 값 ${overrides.size}개 포함</div>`);
  }

  convertReviewSummary.innerHTML = lines.join('');

  const showPreview = !workbench.isVisual();
  convertReviewPreview.hidden = !showPreview;
  convertReviewTableHead.innerHTML = '';
  convertReviewTableBody.innerHTML = '';
  if (showPreview) {
    labelsB.forEach((label) => {
      const th = document.createElement('th');
      th.textContent = label;
      convertReviewTableHead.appendChild(th);
    });
    const shown = Math.min(derived.rowCount, CONVERT_PREVIEW_ROW_LIMIT);
    for (let i = 0; i < shown; i++) {
      const tr = document.createElement('tr');
      headersB.forEach((b) => {
        const td = document.createElement('td');
        td.textContent = displayValue(derived.value(i, b), b);
        tr.appendChild(td);
      });
      convertReviewTableBody.appendChild(tr);
    }
    if (derived.rowCount > shown) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      td.colSpan = headersB.length;
      td.className = 'empty-hint';
      td.textContent = `... 외 ${derived.rowCount - shown}행`;
      tr.appendChild(td);
      convertReviewTableBody.appendChild(tr);
    }
  }
  return confirmModal(convertReviewModal, convertReviewProceedBtn, convertReviewCancelBtn);
}

/** Date를 엑셀 날짜 일련번호로. 시간대 계산 없이 화면에 보이는 연/월/일/시/분/초 그대로 옮긴다 */
function toExcelDateSerial(d) {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
  return (utc - Date.UTC(1899, 11, 30)) / 86400000;
}

/**
 * [Load] 변환된 데이터를 서식이 적용된 엑셀 파일로 내보낸다 (exceljs 사용).
 * - 헤더 행: 배경색 + 흰 글자 + 굵게 + 테두리, 첫 행 고정(스크롤해도 헤더 유지)
 * - 열 너비: 헤더/값의 실제 표시 폭에 맞춰 자동 조정 → "###" 잘림 방지
 * - 숫자 열: 빈 값을 뺀 모든 값이 숫자인 열은 실제 숫자로 저장하고 천단위 구분(#,##0)
 *   서식과 우측 정렬을 적용한다. 원본이 엑셀이면 원본에서 숫자였던 값만 숫자로 본다 —
 *   원본 작성자가 일부러 "텍스트"로 넣은 사번·계좌번호 같은 숫자 글자를 바꾸지 않기
 *   위함. 값의 종류 정보가 없는 원본(CSV, 한글 문서 표: numericText)은 "안전한 숫자
 *   글자"(isSafeNumericString)도 숫자로 바꾼다. columnKinds를 넘기면(매핑 작업대) 그
 *   판정을 그대로 쓰고, 숫자 열 안에서 글자로 고친 칸은 그 칸만 글자로 쓴다.
 * - 날짜: 엑셀 날짜 일련번호 + "yyyy-mm-dd"(시각이 있으면 "yyyy-mm-dd hh:mm:ss") 서식.
 *   JS Date를 exceljs에 그대로 넘기면 UTC 기준으로 저장되어 한국 시간에서는 하루 전
 *   날짜가 되고 표시 형식도 미국식(mm-dd-yy)이 되므로 직접 일련번호로 바꾼다.
 * - numberFormats: 원본 컬럼의 숫자 표시 형식(%, 통화 등). 있으면 기본 천단위 서식 대신 쓴다.
 * headers는 값을 찾을 컬럼키, labels는 헤더 행에 보일 이름이다.
 */
async function buildStyledExcelBuffer(headers, labels, rows, sheetName, { numericText = false, numberFormats = {}, columnKinds = null } = {}) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName || 'Converted');

  worksheet.addRow(labels);
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

  const isNumericValue = (v) => isNumericCell(v, numericText);
  const columnIsNumeric = columnKinds
    ? columnKinds.map((k) => k.kind === 'number')
    : headers.map((h) => {
        let sawAny = false;
        for (const row of rows) {
          const v = row[h];
          if (isEmptyValue(v)) continue;
          sawAny = true;
          if (!isNumericValue(v)) return false;
        }
        return sawAny;
      });
  const columnHasDecimal = headers.map(
    (h, i) => columnIsNumeric[i] && rows.some((row) => isNumericValue(row[h]) && !Number.isInteger(Number(row[h])))
  );

  rows.forEach((row) => {
    const values = headers.map((h, i) => {
      const raw = row[h];
      if (isEmptyValue(raw)) return '';
      if (raw instanceof Date) return toExcelDateSerial(raw);
      return columnIsNumeric[i] && isNumericValue(raw) ? Number(raw) : raw;
    });
    const excelRow = worksheet.addRow(values);
    excelRow.eachCell((cell, colNumber) => {
      const i = colNumber - 1;
      const raw = row[headers[i]];
      if (raw instanceof Date) {
        const hasTime = raw.getHours() !== 0 || raw.getMinutes() !== 0 || raw.getSeconds() !== 0;
        cell.numFmt = hasTime ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd';
        cell.alignment = { horizontal: 'center' };
      } else if (columnIsNumeric[i] && isNumericValue(raw)) {
        cell.numFmt = numberFormats[headers[i]] || (columnHasDecimal[i] ? '#,##0.######' : '#,##0');
        cell.alignment = { horizontal: 'right' };
      } else {
        cell.alignment = { horizontal: 'left' };
      }
    });
  });

  headers.forEach((h, i) => {
    let maxWidth = visualWidth(labels[i]);
    rows.forEach((row) => {
      const v = row[h];
      if (isEmptyValue(v)) return;
      maxWidth = Math.max(maxWidth, visualWidth(formatPreviewValue(v)));
    });
    worksheet.getColumn(i + 1).width = Math.min(40, Math.max(8, Math.ceil(maxWidth) + 2));
  });

  return workbook.xlsx.writeBuffer();
}

/** 3단계 카드 설명: 몇 행을 저장하고 고친 값이 몇 개 들어가는지 */
function updateConvertDesc() {
  if (!convertDesc) return;
  if (!source) {
    convertDesc.textContent = '매핑 작업대의 결과 미리보기와 똑같은 표를 서식이 적용된 엑셀(.xlsx)로 저장합니다.';
    return;
  }
  const edited = overrides.size > 0 ? ` 직접 고친 값 ${overrides.size}개가 들어갑니다.` : '';
  convertDesc.textContent = workbench.isVisual()
    ? `미리보기 그대로 ${source.rows.length}행을 저장합니다.${edited}`
    : `매핑에 따라 ${source.rows.length}행을 저장합니다. 저장 전에 검증 결과와 앞부분 미리보기를 확인할 수 있습니다.${edited}`;
}
workbench.onRender(updateConvertDesc);
workbench.render();

document.getElementById('btnConvert').addEventListener('click', async () => {
  try {
    if (!fileAPath || !source || headersA.length === 0) {
      throw new Error('원본(A) 파일 분석이 필요합니다.');
    }
    if (headersB.length === 0) {
      throw new Error('타겟(B) 파일 분석이 필요합니다.');
    }

    const unmappedColumns = headersB.filter((b) => !mapping[b]);
    const mappedCount = headersB.length - unmappedColumns.length;
    if (mappedCount === 0 && overrides.size === 0) {
      throw new Error('매핑된 컬럼이 하나도 없습니다. 매핑 작업대를 확인하세요.');
    }
    if (source.rows.length === 0) {
      throw new Error('원본(A) 파일에 변환할 데이터 행이 없습니다.');
    }
    if (unmappedColumns.length > 0) {
      log(`매핑되지 않은 컬럼 ${unmappedColumns.length}개는 빈 값으로 채워집니다(직접 고친 칸 제외): ${unmappedColumns.join(', ')}`, 'warn');
    }

    // [Transform] 미리보기와 같은 파생 데이터(원본 값 + 직접 고친 값)를 그대로 쓴다
    const derived = getDerived();
    log(`총 ${derived.rowCount}행이 변환되었습니다.`, 'ok');

    const missingTotal = Object.values(derived.missingByColumn).reduce((sum, n) => sum + n, 0);
    if (missingTotal > 0) {
      log(`필수 컬럼 값이 비어 있는 행이 있습니다 (총 ${missingTotal}건). 확인 창에서 자세한 내용을 볼 수 있습니다.`, 'warn');
    }
    if (derived.duplicateCount > 0) {
      log(`모든 컬럼 값이 동일한 중복 행이 ${derived.duplicateCount}건 있습니다.`, 'warn');
    }

    const proceed = await showConvertReviewModal(derived);
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

    const numberFormats = {};
    headersB.forEach((bCol) => {
      const aCol = mapping[bCol];
      if (aCol && source.numberFormats[aCol]) numberFormats[bCol] = source.numberFormats[aCol];
    });

    // [Load] 헤더 서식·열 너비·숫자 서식을 적용하여 완성된 엑셀 파일로 저장한다
    const outputBuffer = await buildStyledExcelBuffer(headersB, labelsB, derived.getRows(), sheetBName, {
      numericText: source.numericText,
      numberFormats,
      columnKinds: derived.kinds,
    });
    safeWriteFileSync(savePath, outputBuffer);

    log(`변환된 파일이 저장되었습니다: ${savePath}${overrides.size > 0 ? ` (직접 고친 값 ${overrides.size}개 포함)` : ''}`, 'ok');
    alert('✅ 변환이 완료되었습니다.');
  } catch (err) {
    notifyError('데이터 변환 중 오류가 발생했습니다', err);
  }
});
