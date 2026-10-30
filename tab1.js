// tab1.js
// TAB 1. 문서 양식 변환기: 구조 분석, 매핑 작업대(상태 모델·실행 취소·결과 미리보기·값 수정),
// 프리셋, 서식 적용 저장(ETL). 화면 그리기와 조작은 mapping-workbench.js가 맡고,
// 상태는 이 파일만 가진다(작업대는 api를 통해 읽고 바꾼다).

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const excelWriter = require('./excel-writer.js');
const outputCheck = require('./output-check.js');
const verifyWorker = require('./verify-worker.js');
const childProcess = require('child_process');
const XLSX = require('xlsx');
const common = require('./common.js');
const { createWorkbench } = require('./mapping-workbench.js');
const mergeSources = require('./merge-sources.js');

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
  describeRecognition,
  readSheetTable,
  readWorkbookSmart,
  getSynonymGroups,
  autoMatchAll,
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
// 파일별 읽기 설정: 헤더 행(0부터, null이면 자동)과 글자 인코딩(CSV). 파일·시트를 바꾸면 자동으로 돌아간다
const readOptions = { A: { headerRowIndex: null, encoding: 'auto' }, B: { headerRowIndex: null, encoding: 'auto' } };

/* -------------------------------------------------------------------------
 * 상태 모델 (탭1의 유일한 진실 원천)
 * [구조 분석] 때 원본을 한 번 읽어 source에 둔다. 변환 때 다시 읽지 않으므로
 * 미리보기와 저장 결과가 어긋나지 않는다. 화면 요소에서 상태를 거꾸로 읽지 않고,
 * 모든 변경은 commit()을 거친다.
 * ------------------------------------------------------------------------- */

let source = null; // readSheetTable(+합치기) 결과 + { samples, valueKinds, emptyColumns, columnOrigin, joinReports }
let headersA = [];
let headersB = []; // 컬럼키 (같은 이름이 둘이면 "비고_1"처럼 구분됨)
let labelsB = []; // 결과 파일 헤더에 쓸 원래 컬럼 이름
let mapping = {}; // { B키: A키 | '' }
let matchReason = {}; // { B키: 'exact'|'synonym'|'fuzzy'|'confirmed'|'manual'|'preset'|'rule'|'none' }
let autoCandidates = {}; // { B키: { list: [A키], note } } 자동 매핑이 모호해서 멈춘 후보와 그 이유
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

/**
 * 날짜로 쓴 글자를 현지 시각 Date로 바꾼다. 아니면 null.
 * 받는 형식: 2026-01-05 · 2026/1/5 · 2026.01.05 · "2026. 1. 5." · 2026년 1월 5일, 뒤에 시각 13:05 또는 13:05:30.
 * 연도가 4자리여야 하고 실제로 있는 날짜여야 한다(2026-02-30은 날짜가 아님).
 */
const TEXT_DATE_PATTERN =
  /^(\d{4})\s*(?:[-/.]\s*|년\s*)(\d{1,2})\s*(?:[-/.]\s*|월\s*)(\d{1,2})\s*(?:\.|일)?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;

function parseTextDate(str) {
  const m = TEXT_DATE_PATTERN.exec(String(str).trim());
  if (!m) return null;
  const [y, mo, d, h, mi, s] = m.slice(1).map((x) => (x === undefined ? 0 : Number(x)));
  const date = new Date(y, mo - 1, d, h, mi, s);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d || h > 23 || mi > 59 || s > 59) return null;
  return date;
}

/** 직접 고친 값을 날짜 열에 넣을 때도 같은 형식을 받는다 */
const parseOverrideDate = parseTextDate;

/**
 * 값이 모두 글자로 들어오는 원본(CSV, 한글 문서 표)에서 날짜 열을 실제 날짜로 바꾼다.
 * 숫자 열 규칙과 같이 열 단위로 판정한다: 빈 값을 뺀 모든 값이 날짜 글자일 때만 그 열 전체를
 * 바꾸고, 하나라도 날짜가 아니면 열 전체를 글자 그대로 둔다(일부 칸만 날짜가 되는 일이 없게).
 * 바꾼 컬럼 이름 목록을 돌려준다.
 */
function convertTextDateColumns(table) {
  const converted = [];
  table.headers.forEach((a) => {
    const parsed = [];
    for (let i = 0; i < table.rows.length; i++) {
      const v = table.rows[i][a];
      if (isEmptyValue(v)) continue;
      if (typeof v !== 'string') return;
      const date = parseTextDate(v);
      if (!date) return;
      parsed.push([i, date]);
    }
    if (parsed.length === 0) return;
    parsed.forEach(([i, date]) => {
      table.rows[i][a] = date;
    });
    converted.push(a);
  });
  return converted;
}

/**
 * 값이 모두 글자로 들어오는 원본에서, 빈 값을 뺀 모든 값이 "안전한 숫자 글자"(isSafeNumericString)인
 * 열을 숫자로 바꾼다. 앞자리 0이 있는 코드("00123")나 다른 글자가 섞인 열은 글자 그대로 둔다 —
 * 엑셀 원본이 숫자였던 칸만 숫자로 저장하는 것과 같은 결과가 되도록 읽을 때 맞춰 둔다.
 */
function convertTextNumberColumns(table) {
  const converted = [];
  table.headers.forEach((a) => {
    let sawAny = false;
    for (let i = 0; i < table.rows.length; i++) {
      const v = table.rows[i][a];
      if (isEmptyValue(v)) continue;
      if (typeof v !== 'string' || !isSafeNumericString(v)) return;
      sawAny = true;
    }
    if (!sawAny) return;
    table.rows.forEach((row) => {
      if (!isEmptyValue(row[a])) row[a] = Number(String(row[a]).trim());
    });
    converted.push(a);
  });
  return converted;
}

/** CSV·한글 문서 표(값이 모두 글자인 원본)이면 숫자·날짜 열을 실제 숫자·날짜로 바꾼다 */
function isTextSource(filePath) {
  return ['.csv', '.hwp', '.hwpx', '.hwpt'].includes(path.extname(filePath).toLowerCase());
}

function normalizeTextSource(table, filePath, label) {
  if (!(table.textSource || isTextSource(filePath))) return;
  const dateColumns = convertTextDateColumns(table);
  convertTextNumberColumns(table);
  if (dateColumns.length > 0 && label) log(`${label}의 날짜 글자 컬럼을 날짜로 읽었습니다: ${dateColumns.join(', ')}`);
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
function isNumericCell(v) {
  return typeof v === 'number' && Number.isFinite(v);
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
    if (isNumericCell(v)) {
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
 * - number: 빈 값을 뺀 모든 값이 숫자(엑셀 원본은 원본에서 숫자였던 값만, CSV·한글 문서 표는
 *   읽을 때 숫자로 바꾼 값) — buildStyledExcelBuffer의 숫자 열 규칙과 같다.
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
        if (isNumericCell(v) && !Number.isInteger(Number(v))) decimal += 1;
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
  if (kind && kind.kind === 'number' && isNumericCell(v)) {
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
  return derived && ci >= 0 && derived.kinds[ci].kind === 'number' && isNumericCell(v) ? 'num' : '';
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
      sourceName: fileAPath ? [fileAPath, ...((source && source.joinReports) || []).map((r) => r.label)].map((p, i) => (i === 0 ? path.basename(p) : p)).join(' + ') : '',
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

// 시트·헤더 행 선택 <select>와 인식 결과 클릭이 드롭존의 파일선택 클릭으로 전파되지 않도록 차단
[sheetRowA, sheetRowB, ...['A', 'B'].flatMap((w) => [document.getElementById(`readRow${w}`), document.getElementById(`recog${w}`)])].forEach((row) => {
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
  const resultEl = document.getElementById('convertResult');
  if (resultEl) resultEl.hidden = true;
  if (mergeReportEl) {
    mergeReportEl.hidden = true;
    mergeReportEl.innerHTML = '';
  }
  splitBy = '';
  renderSplitOptions();
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
  readOptions[which].headerRowIndex = null;
  resetAnalysisState();
  renderRecognition(which);
  if (which === 'A') refreshExtraSources({ resetKeys: true });
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
    readOptions[which] = { headerRowIndex: null, encoding: 'auto' };
    log(`${which === 'A' ? '원본(A)' : '타겟(B)'} 파일 선택됨: ${p}`);
    const readStart = performance.now();
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
    renderRecognition(which);
    // 파일 읽기(파싱 + 표 만들기)는 여기서 한 번 하고 [구조 분석]은 읽어 둔 워크북을 쓴다 — 시간도 여기서 잰다
    log(`${which === 'A' ? '원본(A)' : '타겟(B)'} 파일 읽기 ${Math.round(performance.now() - readStart)}ms`);
    if (which === 'A') refreshExtraSources({ resetKeys: true });
  });
}

/** 원본(A)·타겟(B)을 지금 읽기 설정(시트·헤더 행·인코딩)으로 읽는다 */
function readTableFor(which, options = {}) {
  const filePath = which === 'A' ? fileAPath : fileBPath;
  const sheetName = which === 'A' ? sheetAName : sheetBName;
  return readSheetTable(filePath, sheetName, { ...readOptions[which], ...options });
}

/**
 * ① 인식 결과: 헤더로 고른 행, 건너뛴 위쪽 행, 이름 없는 열, 인코딩, 확장자와 다른 내용을 보여 주고
 * 헤더 행(맨 위 10행 중)과 CSV 인코딩을 직접 고를 수 있게 한다.
 */
function renderRecognition(which) {
  const readRow = document.getElementById(`readRow${which}`);
  const list = document.getElementById(`recog${which}`);
  const headerSelect = document.getElementById(`headerSelect${which}`);
  const encodingSelect = document.getElementById(`encodingSelect${which}`);
  const encodingLabel = readRow.querySelector('.encoding-label');
  const filePath = which === 'A' ? fileAPath : fileBPath;
  list.innerHTML = '';
  if (!filePath) {
    readRow.hidden = true;
    list.hidden = true;
    return;
  }
  const addItem = ({ text, warn }) => {
    const li = document.createElement('li');
    li.textContent = text;
    if (warn) li.className = 'warn';
    list.appendChild(li);
  };
  let table;
  try {
    table = readTableFor(which);
  } catch (err) {
    readRow.hidden = true;
    list.hidden = false;
    addItem({ text: `읽지 못했습니다: ${friendlyErrorMessage(err)}`, warn: true });
    return;
  }
  const { report, fileInfo } = table;
  headerSelect.innerHTML = '';
  headerSelect.appendChild(new Option(`자동 (${report.autoHeaderRowNumber}행)`, ''));
  report.headerCandidates.forEach((c) => {
    if (c.filled === 0) return;
    const opt = new Option(`${c.rowNumber}행: ${c.preview}`, String(c.index));
    headerSelect.appendChild(opt);
  });
  headerSelect.value = readOptions[which].headerRowIndex === null ? '' : String(readOptions[which].headerRowIndex);
  const isText = !!fileInfo.encoding;
  encodingSelect.hidden = !isText;
  encodingLabel.hidden = !isText;
  encodingSelect.value = readOptions[which].encoding;
  readRow.hidden = false;
  list.hidden = false;
  describeRecognition(report, fileInfo, { includeRows: which === 'A' }).forEach(addItem);
}

/** 헤더 행·인코딩을 바꾸면 분석 결과와 어긋나므로(고친 값이 있으면 확인 후) 초기화한다 */
async function changeReadOption(which, key, value) {
  if (readOptions[which][key] === value) return;
  if (!(await confirmDiscardOverrides(key === 'encoding' ? '인코딩을 바꿀까요?' : '헤더 행을 바꿀까요?'))) {
    renderRecognition(which);
    return;
  }
  readOptions[which][key] = value;
  if (key === 'encoding') readOptions[which].headerRowIndex = null;
  const hadAnalysis = !!source;
  resetAnalysisState();
  renderRecognition(which);
  if (which === 'A') refreshExtraSources({ resetKeys: true });
  const label = which === 'A' ? '원본(A)' : '타겟(B)';
  log(`${label} 읽기 설정을 바꿨습니다(${key === 'encoding' ? `인코딩: ${value}` : value === null ? '헤더 행: 자동' : `헤더 행: ${value + 1}번째 행`}).${hadAnalysis ? ' [구조 분석]을 다시 실행하세요.' : ''}`, hadAnalysis ? 'warn' : 'info');
}

['A', 'B'].forEach((which) => {
  document.getElementById(`headerSelect${which}`).addEventListener('change', (e) => {
    changeReadOption(which, 'headerRowIndex', e.target.value === '' ? null : Number(e.target.value));
  });
  document.getElementById(`encodingSelect${which}`).addEventListener('change', (e) => {
    changeReadOption(which, 'encoding', e.target.value);
  });
});

bindTab1File('A', dropA, fileNameA, sheetRowA, sheetSelectA);
bindTab1File('B', dropB, fileNameB, sheetRowB, sheetSelectB);

/* -------------------------------------------------------------------------
 * 원본 합치기: 첫 원본(A)에 다른 파일의 컬럼을 기준 컬럼(예: 사번)으로 이어 붙인다.
 * 기준 컬럼은 값이 겹치는 정도·고유도·이름·코드 모양으로 자동 추천하고(merge-sources.js),
 * 사용자가 바꿀 수 있다. 실제로 합치는 것은 [구조 분석] 때다.
 * ------------------------------------------------------------------------- */

const extraSourceList = document.getElementById('extraSourceList');
const mergeReportEl = document.getElementById('mergeReport');
// [{ id, path, label, sheetName, sheetNames, sheetUnit, headers, table, keyMain, keyOther, mode, dupPolicy, ignoreZeros, suggestions, recognition, error }]
let extraSources = [];
let extraSourceSeq = 0;

/** 파일 이름(확장자 제외)으로 짧은 라벨. 겹치면 뒤에 번호 */
function makeSourceLabel(filePath) {
  const base = path.basename(filePath, path.extname(filePath)) || '추가 원본';
  let label = base;
  for (let n = 2; extraSources.some((s) => s.label === label); n++) label = `${base}${n}`;
  return label;
}

/** 첫 원본(A)의 표(글자 숫자·날짜 정리 후). 아직 고르지 않았으면 null */
function readMainTableForMerge() {
  if (!fileAPath) return null;
  try {
    const table = readTableFor('A');
    normalizeTextSource(table, fileAPath);
    return table;
  } catch (err) {
    return null;
  }
}

/** 추가 원본 하나를 (다시) 읽고, 첫 원본과 이을 기준 컬럼을 추천해 비어 있으면 채운다 */
function loadExtraSource(es, mainTable, { resetKeys = false } = {}) {
  es.error = '';
  try {
    const table = readSheetTable(es.path, es.sheetName);
    normalizeTextSource(table, es.path);
    es.headers = table.headers;
    es.table = table;
    es.recognition = describeRecognition(table.report, table.fileInfo).filter((item) => item.warn);
    if (mainTable) {
      es.suggestions = mergeSources.suggestJoinKeys(mainTable, table, { synonymGroups: getSynonymGroups() });
      const keysValid = es.keyMain && mainTable.headers.includes(es.keyMain) && es.headers.includes(es.keyOther);
      if (resetKeys || !keysValid) {
        const best = es.suggestions[0];
        es.keyMain = best ? best.main : '';
        es.keyOther = best ? best.other : '';
      }
    } else {
      es.suggestions = [];
    }
  } catch (err) {
    es.headers = [];
    es.suggestions = [];
    es.error = friendlyErrorMessage(err);
  }
}

function refreshExtraSources({ resetKeys = false } = {}) {
  if (extraSources.length === 0) return;
  const mainTable = readMainTableForMerge();
  extraSources.forEach((es) => loadExtraSource(es, mainTable, { resetKeys }));
  renderExtraSources();
}

function renderExtraSources() {
  extraSourceList.innerHTML = '';
  const mainTable = fileAPath ? readMainTableForMerge() : null;
  const mainHeaders = mainTable ? mainTable.headers : [];
  extraSources.forEach((es) => {
    const item = document.createElement('div');
    item.className = 'extra-source';
    item.dataset.id = String(es.id);

    const title = document.createElement('div');
    title.className = 'extra-source-title';
    const name = document.createElement('b');
    name.textContent = path.basename(es.path);
    title.appendChild(name);
    if (es.sheetNames.length > 1) {
      const sheetSelect = document.createElement('select');
      sheetSelect.className = 'extra-sheet';
      sheetSelect.setAttribute('aria-label', `${path.basename(es.path)} ${es.sheetUnit} 선택`);
      es.sheetNames.forEach((n) => sheetSelect.appendChild(new Option(n, n)));
      sheetSelect.value = es.sheetName;
      sheetSelect.addEventListener('change', () => changeExtraSource(es, () => {
        es.sheetName = sheetSelect.value;
        loadExtraSource(es, readMainTableForMerge(), { resetKeys: true });
      }));
      title.appendChild(sheetSelect);
    }
    const spacer = document.createElement('span');
    spacer.className = 'wb-spacer';
    title.appendChild(spacer);
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'btn ghost small extra-remove';
    remove.textContent = '빼기';
    remove.setAttribute('aria-label', `${path.basename(es.path)} 합치기에서 빼기`);
    remove.addEventListener('click', () => changeExtraSource(es, () => {
      extraSources = extraSources.filter((x) => x !== es);
    }));
    title.appendChild(remove);
    item.appendChild(title);

    if (es.error) {
      const err = document.createElement('div');
      err.className = 'extra-source-error';
      err.textContent = `읽지 못했습니다: ${es.error}`;
      item.appendChild(err);
      extraSourceList.appendChild(item);
      return;
    }

    const keyRow = document.createElement('div');
    keyRow.className = 'extra-source-keys';
    const keySelect = (options, value, label, onChange) => {
      const sel = document.createElement('select');
      sel.setAttribute('aria-label', label);
      sel.appendChild(new Option('-- 기준 컬럼 --', ''));
      options.forEach((h) => sel.appendChild(new Option(h, h)));
      sel.value = value || '';
      sel.addEventListener('change', () => changeExtraSource(es, () => onChange(sel.value)));
      return sel;
    };
    const lab = (text) => {
      const s = document.createElement('span');
      s.className = 'extra-label';
      s.textContent = text;
      return s;
    };
    keyRow.append(
      lab('기준'),
      keySelect(mainHeaders, es.keyMain, '원본(A)의 기준 컬럼', (v) => {
        es.keyMain = v;
      }),
      lab('='),
      keySelect(es.headers, es.keyOther, `${es.label}의 기준 컬럼`, (v) => {
        es.keyOther = v;
      })
    );
    const modeSelect = document.createElement('select');
    modeSelect.setAttribute('aria-label', '합치는 방식');
    modeSelect.appendChild(new Option('원본(A)에 있는 행만', 'left'));
    modeSelect.appendChild(new Option('양쪽 행 모두', 'outer'));
    modeSelect.value = es.mode;
    modeSelect.addEventListener('change', () => changeExtraSource(es, () => {
      es.mode = modeSelect.value;
    }));
    keyRow.append(lab('행'), modeSelect);
    item.appendChild(keyRow);

    // 기준값이 겹칠 때(한 사번이 두 행) 어떻게 할지, 앞자리 0을 무시할지 — 둘 다 사용자가 고른다
    const optionRow = document.createElement('div');
    optionRow.className = 'extra-source-keys';
    const dupSelect = document.createElement('select');
    dupSelect.setAttribute('aria-label', '기준값이 겹칠 때');
    dupSelect.appendChild(new Option('첫 행만 잇기', 'first'));
    dupSelect.appendChild(new Option('모두 잇기(행이 늘어남)', 'all'));
    dupSelect.value = es.dupPolicy || 'first';
    dupSelect.addEventListener('change', () => changeExtraSource(es, () => {
      es.dupPolicy = dupSelect.value;
    }));
    const zeroLabel = document.createElement('label');
    zeroLabel.className = 'extra-zero';
    const zeroBox = document.createElement('input');
    zeroBox.type = 'checkbox';
    zeroBox.checked = !!es.ignoreZeros;
    zeroBox.addEventListener('change', () => changeExtraSource(es, () => {
      es.ignoreZeros = zeroBox.checked;
    }));
    zeroLabel.append(zeroBox, document.createTextNode(' 앞자리 0 무시(00123 = 123)'));
    optionRow.append(lab('겹치는 기준값'), dupSelect, zeroLabel);
    item.appendChild(optionRow);

    const hint = document.createElement('div');
    hint.className = 'extra-source-hint';
    const chosen = es.suggestions.find((s) => s.main === es.keyMain && s.other === es.keyOther);
    if (!fileAPath) hint.textContent = '원본(A)을 먼저 고르면 기준 컬럼을 추천합니다.';
    else if (chosen) hint.textContent = `${chosen === es.suggestions[0] ? '추천' : '후보'}: ${chosen.reasons.join(' · ')}`;
    else if (es.keyMain && es.keyOther) hint.textContent = '직접 고른 기준입니다. 값이 얼마나 겹치는지는 [구조 분석] 결과에서 확인하세요.';
    else {
      hint.textContent = '값이 겹치는 기준 컬럼을 찾지 못했습니다. 직접 골라 주세요.';
      hint.classList.add('warn');
    }
    item.appendChild(hint);

    // 앞자리 0 무시를 고르기 전에: 정확히 비교할 때와 무시할 때 연결되는 행 수, 같아지는 서로 다른 값
    if (mainTable && es.table && es.keyMain && es.keyOther && mainTable.headers.includes(es.keyMain) && es.headers.includes(es.keyOther)) {
      const info = mergeSources.analyzeKeyMatch(mainTable, es.keyMain, es.table, es.keyOther);
      const parts = [`정확히 비교하면 원본(A) ${info.exactMatched}행 연결`];
      if (info.looseMatched !== info.exactMatched || es.ignoreZeros) parts.push(`앞자리 0을 무시하면 ${info.looseMatched}행 연결`);
      if (info.collisions) parts.push(`무시하면 서로 다른 값 ${info.collisions}쌍이 같아짐(예: ${info.collisionExamples.join(', ')})`);
      const counts = document.createElement('div');
      counts.className = `extra-source-hint extra-match-counts${(es.ignoreZeros && info.collisions) || (!es.ignoreZeros && info.exactMatched === 0) ? ' warn' : ''}`;
      counts.textContent = parts.join(' · ');
      item.appendChild(counts);
    }
    (es.recognition || []).forEach(({ text }) => {
      const note = document.createElement('div');
      note.className = 'extra-source-hint warn';
      note.textContent = `⚠ ${text}`;
      item.appendChild(note);
    });
    extraSourceList.appendChild(item);
  });
}

/** 합치기 설정을 바꾼다: 분석 결과가 새 설정과 어긋나지 않도록(고친 값이 있으면 확인 후) 초기화 */
async function changeExtraSource(es, apply) {
  if (source && !(await confirmDiscardOverrides('원본 합치기 설정을 바꿀까요?'))) {
    renderExtraSources();
    return;
  }
  apply();
  if (source) {
    resetAnalysisState();
    log('원본 합치기 설정이 바뀌었습니다. [구조 분석]을 다시 실행하세요.', 'warn');
  }
  renderExtraSources();
}

document.getElementById('btnAddSource').addEventListener('click', async () => {
  try {
    const paths = await ipcRenderer.invoke('dialog:openFile', { filters: DATA_FILE_FILTERS, multi: false });
    if (!paths || paths.length === 0) return;
    const p = paths[0];
    const workbook = readWorkbookSmart(p);
    const sheetNames = workbook.SheetNames || [];
    const es = {
      id: ++extraSourceSeq,
      path: p,
      label: makeSourceLabel(p),
      sheetNames,
      sheetUnit: workbook.sheetUnitLabel || '시트',
      sheetName: sheetNames.includes(workbook.preferredSheetName) ? workbook.preferredSheetName : sheetNames[0],
      headers: [],
      keyMain: '',
      keyOther: '',
      mode: 'left',
      dupPolicy: 'first',
      ignoreZeros: false,
      suggestions: [],
      error: '',
    };
    await changeExtraSource(es, () => {
      loadExtraSource(es, readMainTableForMerge(), { resetKeys: true });
      extraSources.push(es);
    });
    log(`합칠 원본 추가: ${p}${es.keyMain ? ` (기준 ${es.keyMain} = ${es.keyOther})` : ''}`);
  } catch (err) {
    notifyError('원본 추가 중 오류가 발생했습니다', err);
  }
});

/** [구조 분석] 때: 첫 원본에 추가 원본들을 합친 표. 추가 원본이 없으면 첫 원본 그대로 */
function joinExtraSources(mainTable) {
  if (extraSources.length === 0) return mainTable;
  const others = extraSources.map((es) => {
    if (es.error) throw new Error(`추가 원본 "${path.basename(es.path)}"을(를) 읽지 못했습니다: ${es.error}`);
    if (!es.keyMain || !es.keyOther) throw new Error(`추가 원본 "${path.basename(es.path)}"의 기준 컬럼을 고르세요.`);
    if (!mainTable.headers.includes(es.keyMain)) throw new Error(`원본(A)에 기준 컬럼 "${es.keyMain}"이(가) 없습니다. 합치기 설정을 확인하세요.`);
    const table = readSheetTable(es.path, es.sheetName);
    normalizeTextSource(table, es.path, es.label);
    if (!table.headers.includes(es.keyOther)) throw new Error(`"${path.basename(es.path)}"에 기준 컬럼 "${es.keyOther}"이(가) 없습니다.`);
    return { label: es.label, table, keyMain: es.keyMain, keyOther: es.keyOther, mode: es.mode, dupPolicy: es.dupPolicy || 'first', ignoreLeadingZeros: !!es.ignoreZeros };
  });
  const { table, reports } = mergeSources.joinTables(mainTable, others);
  reports.forEach((r) => log(`원본 합치기 — ${mergeSources.describeJoinReport(r)}`, mergeSources.describeJoinReportDetail(r).warn ? 'warn' : 'ok'));
  return { ...mainTable, ...table, joinReports: reports };
}

function renderMergeReport() {
  mergeReportEl.innerHTML = '';
  const reports = source ? source.joinReports : [];
  mergeReportEl.hidden = reports.length === 0;
  reports.forEach((r) => {
    const detail = mergeSources.describeJoinReportDetail(r);
    const line = document.createElement('div');
    line.className = `merge-report-line${detail.warn ? ' warn' : ''}`;
    line.textContent = detail.text;
    mergeReportEl.appendChild(line);
    // 빠진 행·덧붙인 행은 이유별로 행 번호와 함께
    detail.details.forEach((d) => {
      const sub = document.createElement('div');
      sub.className = `merge-report-detail${d.warn ? ' warn' : ''}`;
      sub.textContent = d.text;
      mergeReportEl.appendChild(sub);
    });
  });
}

/** 원본 표(합친 결과)에 화면용 정보(컬럼별 앞쪽 값, 값 모양, 값이 모두 빈 컬럼)를 붙인다 */
function prepareSource(table) {
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
    columnOrigin: table.columnOrigin || {},
    joinReports: table.joinReports || [],
  };
}

/**
 * 분석 직후 자동 매핑. 타겟 컬럼 순서와 상관없이 같은 이름을 모두 먼저 확정한 뒤 동의어·비슷한 이름을
 * 정한다(common.autoMatchAll). 동의어/비슷한 이름 연결과 모호해서 멈춘 컬럼은 로그로도 알린다.
 */
function runAutoMapping() {
  mapping = {};
  matchReason = {};
  autoCandidates = {};
  const results = autoMatchAll(headersB, headersA);
  headersB.forEach((bCol) => {
    const { match, reason, candidates, note } = results[bCol];
    if (match) {
      mapping[bCol] = match;
      matchReason[bCol] = reason;
      if (reason === 'synonym') {
        log(`"${bCol}" 컬럼을 동의어 사전에 따라 "${match}" 컬럼에 자동 매핑했습니다.`);
      } else if (reason === 'fuzzy') {
        log(`"${bCol}" 컬럼을 이름이 비슷한 "${match}" 컬럼에 자동 매핑했습니다. 맞는지 확인하세요.`, 'warn');
      }
    } else {
      mapping[bCol] = '';
      matchReason[bCol] = 'none';
      if (candidates && candidates.length > 0) {
        autoCandidates[bCol] = { list: candidates, note };
        log(`"${bCol}" 컬럼은 자동 연결하지 않았습니다(${note}: ${candidates.join(', ')}). 직접 고르세요.`, 'warn');
      }
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
    const analyzeStart = performance.now();
    const resultA = readTableFor('A');
    normalizeTextSource(resultA, fileAPath, '원본(A)');
    const resultB = readTableFor('B');
    [
      ['원본(A)', resultA, true],
      ['타겟(B)', resultB, false],
    ].forEach(([label, t, includeRows]) => {
      describeRecognition(t.report, t.fileInfo, { includeRows })
        .filter((item) => item.warn)
        .forEach((item) => log(`${label} 파일: ${item.text}`, 'warn'));
    });
    renderRecognition('A');
    renderRecognition('B');
    [
      ['원본(A)', resultA],
      ['타겟(B)', resultB],
    ].forEach(([label, { headers, labels }]) => {
      const renamed = headers.filter((k, i) => k !== labels[i]);
      if (renamed.length > 0) log(`${label} 파일에 이름이 같은 컬럼이 있어 ${renamed.join(', ')}(으)로 구분했습니다.`, 'warn');
    });

    const combined = joinExtraSources(resultA);

    resetAnalysisState();
    source = prepareSource(combined);
    headersA = combined.headers;
    renderMergeReport();
    headersB = resultB.headers;
    labelsB = resultB.labels;
    log(`원본(A) 컬럼 ${headersA.length}개${extraSources.length ? `(추가 원본 ${extraSources.length}개 합침)` : ''}, 타겟(B) 컬럼 ${headersB.length}개 발견`, 'ok');

    runAutoMapping();
    invalidate();
    workbench.onAnalyzed();
    workbench.render();
    renderSplitOptions({ recompute: true });
    analyzeStatus.textContent = '분석 완료';
    analyzeStatus.className = 'pill ok';
    log(`구조 분석 ${Math.round(performance.now() - analyzeStart)}ms`);
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
      splitBy: splitBy || undefined,
      splitIncludeAll: splitBy ? splitIncludeAll.checked : undefined,
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
    applyLoadedSplit(payload);
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

/** 프리셋·규칙 파일에 시트 나누기 설정이 있으면 적용(이 양식에 그 컬럼이 있을 때만) */
function applyLoadedSplit(payload) {
  if (!payload || typeof payload.splitBy !== 'string') return;
  if (!headersB.includes(payload.splitBy)) {
    log(`저장된 시트 나누기 컬럼 "${payload.splitBy}"이(가) 이 양식에 없어 한 시트로 저장합니다.`, 'warn');
    return;
  }
  splitBy = payload.splitBy;
  splitIncludeAll.checked = !!payload.splitIncludeAll;
  renderSplitOptions();
}

async function applyPresetById(id) {
  const preset = presets.find((p) => p.id === id);
  if (!preset) return false;
  const { ok, appliedCount, skippedCount } = await applyLoadedMapping(`프리셋 "${preset.name}" 적용`, 'preset', preset.mapping || {}, preset.requiredColumns);
  if (!ok) return false;
  applyLoadedSplit(preset);
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
      splitBy: splitBy || undefined,
      splitIncludeAll: splitBy ? splitIncludeAll.checked : undefined,
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
/**
 * 원본에서 결과까지 행 수가 어떻게 바뀌었는지와 그 이유: [{ text, warn }]
 * 원본(A)을 읽을 때 뺀 행(헤더 위·완전히 빈 행) → 합치기로 덧붙이거나 늘어난 행·빠진 추가 원본 행 → 결과 행 수
 */
function describeRowFlow(rowCount) {
  const items = [];
  const report = source.report;
  if (report) {
    const removed = [];
    if (report.skippedRows.length) removed.push(`헤더 위 ${report.skippedRows.length}행`);
    if (report.emptyRowsRemoved) removed.push(`완전히 빈 행 ${report.emptyRowsRemoved}개`);
    items.push({ text: `원본(A) 데이터 ${report.dataRowCount}행${removed.length ? ` (읽을 때 뺀 것: ${removed.join(', ')})` : ''}`, warn: false });
  }
  (source.joinReports || []).forEach((r) => {
    const parts = [];
    if (r.addedRows) parts.push(`${r.addedRows}행 덧붙임`);
    if (r.expandedRows) parts.push(`겹친 기준값으로 ${r.expandedRows}행 늘어남`);
    if (r.excludedRows) parts.push(`${r.label} 행 ${r.excludedRows}개는 들어가지 않음`);
    if (parts.length) items.push({ text: `합치기 ${r.label}: ${parts.join(', ')}`, warn: r.excludedRows > 0 });
  });
  items.push({ text: `결과 ${rowCount}행`, warn: false });
  return items;
}

function showConvertReviewModal(derived, sheets) {
  const lines = [];
  lines.push(`<div class="review-line ok">총 <b>${derived.rowCount}</b>행을 저장합니다.</div>`);
  const flow = describeRowFlow(derived.rowCount);
  if (flow.length > 1) lines.push(`<div class="review-line ${flow.some((f) => f.warn) ? 'warn' : 'info'}">행 수: ${xmlEscape(flow.map((f) => f.text).join(' → '))}</div>`);
  if (sheets) {
    const parts = sheets.filter((s) => !s.isAll).map((s) => `${s.name} ${s.rows.length}행`);
    const shown = parts.slice(0, 8).join(' · ') + (parts.length > 8 ? ` 외 ${parts.length - 8}개` : '');
    lines.push(`<div class="review-line info">${xmlEscape(labelOfB(splitBy))} 값마다 시트 ${parts.length}개로 나눠 저장${sheets.some((s) => s.isAll) ? '(맨 앞에 전체 시트 포함)' : ''}: ${xmlEscape(shown)}</div>`);
    if (parts.length > SPLIT_WARN_SHEETS) lines.push(`<div class="review-line warn">시트가 ${parts.length}개로 많습니다. 나눌 컬럼이 맞는지 확인하세요.</div>`);
  }
  (source.joinReports || []).forEach((r) => {
    lines.push(`<div class="review-line ${mergeSources.describeJoinReportDetail(r).warn ? 'warn' : 'ok'}">원본 합치기 — ${xmlEscape(mergeSources.describeJoinReport(r))}</div>`);
  });

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



/* -------------------------------------------------------------------------
 * 결과 시트 나누기: 타겟 컬럼 하나의 값마다 시트를 따로 만든다(예: 부서명 → 영업팀·개발팀 시트).
 * 나눌 만한 컬럼(값 종류가 적고 대부분 채워진 컬럼, 이름이 부서·지점·구분 류)을 추천하지만
 * 기본은 한 시트로 저장이다.
 * ------------------------------------------------------------------------- */

const splitSelect = document.getElementById('splitSelect');
const splitIncludeAll = document.getElementById('splitIncludeAll');
const splitHint = document.getElementById('splitHint');
let splitBy = '';
let splitSuggestions = [];
const SPLIT_WARN_SHEETS = 30;

/** 시트 나누기 묶음 기준: 저장될 엑셀에 보이는 글자 */
function splitKeyOf(i, b) {
  const derived = getDerived();
  return displayValue(derived.value(i, b), b);
}

function computeSplitSuggestions() {
  const derived = getDerived();
  if (!derived) return [];
  return mergeSources.suggestSplitColumns(headersB, derived.rowCount, (i, b) => valueToken(derived.value(i, b)));
}

/** 나누기 선택 목록을 다시 그린다. recompute면 추천을 다시 계산(값이 바뀌었을 수 있음) */
function renderSplitOptions({ recompute = false } = {}) {
  if (!source) {
    splitBy = '';
    splitSuggestions = [];
    splitSelect.innerHTML = '';
    splitSelect.appendChild(new Option('한 시트로 저장', ''));
    splitSelect.disabled = true;
    splitIncludeAll.disabled = true;
    splitIncludeAll.checked = false;
    splitHint.textContent = '';
    return;
  }
  if (recompute) splitSuggestions = computeSplitSuggestions();
  if (splitBy && !headersB.includes(splitBy)) splitBy = '';
  splitSelect.innerHTML = '';
  splitSelect.appendChild(new Option('한 시트로 저장', ''));
  const suggested = new Set(splitSuggestions.map((s) => s.column));
  if (splitSuggestions.length > 0) {
    const group = document.createElement('optgroup');
    group.label = '추천';
    splitSuggestions.forEach((s) => group.appendChild(new Option(`${labelOfB(s.column)} 값마다 시트 나누기 (${s.reasons.join(' · ')})`, s.column)));
    splitSelect.appendChild(group);
  }
  const rest = headersB.filter((b) => !suggested.has(b));
  if (rest.length > 0) {
    const group = document.createElement('optgroup');
    group.label = '다른 컬럼';
    rest.forEach((b) => group.appendChild(new Option(`${labelOfB(b)} 값마다 시트 나누기`, b)));
    splitSelect.appendChild(group);
  }
  splitSelect.value = splitBy;
  splitSelect.disabled = false;
  splitIncludeAll.disabled = !splitBy;
  if (splitBy) {
    const count = mergeSources.groupRowsBy(source.rows.length, (i) => splitKeyOf(i, splitBy)).length;
    splitHint.textContent = `${labelOfB(splitBy)} 값마다 시트 ${count}개${count > SPLIT_WARN_SHEETS ? ' — 시트가 많습니다' : ''}`;
  } else if (splitSuggestions.length > 0) {
    const best = splitSuggestions[0];
    splitHint.textContent = `추천: ${labelOfB(best.column)} 컬럼으로 나누기 (값 ${best.values}종류)`;
  } else {
    splitHint.textContent = '';
  }
}

splitSelect.addEventListener('change', () => {
  splitBy = splitSelect.value;
  if (!splitBy) splitIncludeAll.checked = false;
  renderSplitOptions();
});
// 값을 고친 뒤 목록을 열면 추천을 새로 계산한다(매번 그릴 때 계산하면 큰 표에서 무겁다)
splitSelect.addEventListener('mousedown', () => renderSplitOptions({ recompute: true }));
splitSelect.addEventListener('focus', () => renderSplitOptions({ recompute: true }));

/** 저장할 시트 목록. 나누지 않으면 null */
function buildSplitSheets(derived) {
  if (!splitBy) return null;
  const rows = derived.getRows();
  const used = new Set();
  const sheets = [];
  if (splitIncludeAll.checked) sheets.push({ name: mergeSources.sanitizeSheetName(sheetBName || '전체', used), rows, isAll: true });
  mergeSources.groupRowsBy(derived.rowCount, (i) => splitKeyOf(i, splitBy)).forEach((g) => {
    sheets.push({ name: mergeSources.sanitizeSheetName(g.label, used), rows: g.indices.map((i) => rows[i]), blank: g.isBlank });
  });
  return sheets;
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

    const sheets = buildSplitSheets(derived);
    const proceed = await showConvertReviewModal(derived, sheets);
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

    const allRows = derived.getRows();
    const sheetList = sheets || [{ name: sheetBName || 'Converted', rows: allRows }];
    const columnKinds = derived.kinds.map((k) => ({ kind: k.kind }));
    const formatOf = excelWriter.makeCellFormatter(headersB, allRows, { numberFormats, columnKinds });
    const outcome = await saveAndVerify(savePath, { headers: headersB, labels: labelsB, sheets: sheetList, formatOf, numberFormats, columnKinds });
    if (!outcome) return; // 취소

    log(`변환된 파일이 저장되었습니다: ${savePath}${sheets ? ` (시트 ${sheets.length}개)` : ''}${overrides.size > 0 ? ` (직접 고친 값 ${overrides.size}개 포함)` : ''}`, 'ok');
    const { check } = outcome;
    if (check.ok) log(`저장한 파일을 다시 열어 확인했습니다: 시트 ${check.sheetNames.length}개, 값 ${check.checkedCells}칸 모두 같음`, 'ok');
    else log(`저장한 파일을 다시 열어 보니 ${check.problemCount}곳이 다릅니다. ③단계의 확인 결과를 보세요.`, 'warn');
    renderConvertResult(savePath, derived, sheetList, outcome);
  } catch (err) {
    notifyError('데이터 변환 중 오류가 발생했습니다', err);
  } finally {
    setConvertBusy(false);
  }
});

/* -------------------------------------------------------------------------
 * 저장과 ④ 변환 파일 확인
 * 같은 폴더의 임시 파일에 (스트리밍으로) 쓰고 → 그 파일을 디스크에서 다시 열어 시트·헤더·행 수·칸 값·
 * 표시 형식을 저장하려던 표와 비교한 뒤 → 원래 이름으로 바꾼다. 다시 열 수 없으면(깨진 파일) 원래 이름으로
 * 바꾸지 않아, 이미 있던 같은 이름의 파일을 망가뜨리지 않는다. 값이 다른 곳이 있으면 저장은 하되 목록으로 알린다.
 * ------------------------------------------------------------------------- */

const convertProgressEl = document.getElementById('convertProgress');
const convertProgressBar = document.getElementById('convertProgressBar');
const convertProgressText = document.getElementById('convertProgressText');
const convertResultEl = document.getElementById('convertResult');
const btnConvert = document.getElementById('btnConvert');
let cancelSaveRequested = false;
let lastSavedPath = null;

function setConvertBusy(busy, text = '') {
  btnConvert.disabled = busy;
  convertProgressEl.hidden = !busy;
  if (busy) {
    convertProgressBar.value = 0;
    convertProgressText.textContent = text;
  }
}

function removeQuietly(p) {
  try {
    if (fs.existsSync(p)) fs.unlinkSync(p);
  } catch (e) {
    /* 임시 파일 정리 실패는 결과에 영향 없음 */
  }
}

/**
 * 저장한 파일을 별도 프로세스(verify-worker.js)에서 다시 열어 비교한다 — 큰 파일을 여는 동안 창이 멈추지 않고
 * [취소]로 끝낼 수 있다. 프로세스를 띄우지 못하면 이 화면에서 같은 함수로 확인한다.
 * 반환: 비교 결과 · 취소하면 null
 */
function verifyInChild(job) {
  return new Promise((resolve, reject) => {
    let child;
    try {
      child = childProcess.fork(path.join(__dirname, 'verify-worker.js'), [], {
        env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
        serialization: 'advanced', // Date를 그대로 주고받는다
        stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
      });
    } catch (err) {
      resolve(verifyWorker.verifyFile(job));
      return;
    }
    let settled = false;
    let stderr = '';
    child.stderr.on('data', (d) => {
      stderr += d;
    });
    const finish = (fn) => {
      if (settled) return;
      settled = true;
      clearInterval(cancelPoll);
      fn();
    };
    const cancelPoll = setInterval(() => {
      if (!cancelSaveRequested) return;
      child.kill();
      finish(() => resolve(null));
    }, 100);
    child.once('message', (reply) => finish(() => (reply.ok ? resolve(reply.result) : reject(new Error(reply.message)))));
    child.once('error', (err) => finish(() => reject(err)));
    child.once('exit', (code) => finish(() => reject(new Error(`확인 프로세스가 끝났습니다(코드 ${code}) ${stderr.slice(0, 300)}`))));
    child.send(job);
  });
}

/** 반환: { check, writeMs, verifyMs } · 사용자가 취소하면 null */
async function saveAndVerify(savePath, { headers, labels, sheets, formatOf, numberFormats, columnKinds }) {
  const tempPath = path.join(path.dirname(savePath), `~${path.basename(savePath, path.extname(savePath))}.${process.pid}.tmp.xlsx`);
  cancelSaveRequested = false;
  convertResultEl.hidden = true;
  const total = sheets.reduce((n, s) => n + s.rows.length, 0);
  setConvertBusy(true, `0 / ${total}행 쓰는 중…`);
  const t0 = performance.now();
  try {
    await excelWriter.writeStyledExcel(
      {
        headers,
        labels,
        sheets,
        formatOf,
        onProgress: (done, all) => {
          convertProgressBar.value = all ? Math.round((done / all) * 90) : 90;
          convertProgressText.textContent = `${done} / ${all}행 쓰는 중…`;
        },
        shouldCancel: () => cancelSaveRequested,
      },
      tempPath
    );
  } catch (err) {
    removeQuietly(tempPath);
    if (err.cancelled) {
      log('저장을 취소했습니다. 파일은 만들지 않았습니다.', 'warn');
      return null;
    }
    throw err;
  }
  const writeMs = performance.now() - t0;

  convertProgressBar.value = 92;
  convertProgressText.textContent = '저장한 파일을 다시 열어 확인하는 중…';
  await new Promise((r) => setTimeout(r, 0));
  const t1 = performance.now();
  let check;
  try {
    check = await verifyInChild({ filePath: tempPath, headers, labels, sheets, numberFormats, columnKinds });
  } catch (err) {
    removeQuietly(tempPath);
    throw new Error(`저장한 파일을 다시 열 수 없어 저장하지 않았습니다: ${friendlyErrorMessage(err)}`);
  }
  if (!check) {
    removeQuietly(tempPath);
    log('확인 중에 저장을 취소했습니다. 파일은 만들지 않았습니다.', 'warn');
    return null;
  }
  const verifyMs = performance.now() - t1;
  try {
    fs.renameSync(tempPath, savePath);
  } catch (err) {
    removeQuietly(tempPath);
    err.path = err.path || savePath;
    throw err;
  }
  convertProgressBar.value = 100;
  return { check, writeMs, verifyMs };
}

/** ④ 확인 결과: 저장 경로, 행 수 흐름, 시트·헤더·값 비교, 다른 곳 목록 */
function renderConvertResult(savePath, derived, sheetList, { check, writeMs, verifyMs }) {
  lastSavedPath = savePath;
  document.getElementById('convertResultTitle').textContent = check.ok
    ? '✅ 저장하고 다시 열어 확인했습니다 — 저장하려던 표와 같습니다'
    : `⚠ 저장했지만 다시 열어 보니 ${check.problemCount}곳이 다릅니다`;
  convertResultEl.classList.toggle('warn', !check.ok);
  document.getElementById('convertResultPath').textContent = savePath;
  const list = document.getElementById('convertResultItems');
  list.innerHTML = '';
  const add = ({ text, warn }) => {
    const li = document.createElement('li');
    li.textContent = text;
    if (warn) li.className = 'warn';
    list.appendChild(li);
  };
  add({ text: `행 수: ${describeRowFlow(derived.rowCount).map((f) => f.text).join(' → ')}`, warn: describeRowFlow(derived.rowCount).some((f) => f.warn) });
  const rowsProblem = check.problems.some((p) => p.kind === 'rows' || p.kind === 'sheets');
  add({ text: `시트 ${sheetList.length}개: ${sheetList.map((s) => `${s.name} ${check.rowCounts[s.name] !== undefined ? check.rowCounts[s.name] : 0}행`).join(' · ')}`, warn: rowsProblem });
  add({ text: check.headerOk ? `헤더 ${labelsB.length}개 같음` : '헤더 이름이 다른 곳이 있음', warn: !check.headerOk });
  add({
    text: check.ok
      ? `값 ${check.checkedCells}칸을 다시 읽어 비교: 글자·숫자·날짜·표시 형식 모두 같음`
      : `값 ${check.checkedCells}칸을 다시 읽어 비교: 다른 곳 ${check.problemCount}개${check.problemCount > check.problems.length ? ` (앞 ${check.problems.length}개만 표시)` : ''}`,
    warn: !check.ok,
  });
  check.problems.slice(0, 20).forEach((p) => add({ text: outputCheck.describeProblem(p), warn: true }));
  add({ text: `걸린 시간: 쓰기 ${(writeMs / 1000).toFixed(1)}초 · 다시 열어 확인 ${(verifyMs / 1000).toFixed(1)}초`, warn: false });
  convertResultEl.hidden = false;
  convertResultEl.scrollIntoView({ block: 'nearest' });
}

document.getElementById('btnCancelConvert').addEventListener('click', () => {
  cancelSaveRequested = true;
  convertProgressText.textContent = '취소하는 중…';
});

document.getElementById('btnShowSaved').addEventListener('click', () => {
  if (lastSavedPath) ipcRenderer.invoke('shell:showItemInFolder', lastSavedPath);
});
