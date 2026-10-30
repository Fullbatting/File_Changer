// merge-sources.js
// 탭1 "원본 합치기"와 "결과 시트 나누기"의 판단·계산 로직. 화면이나 Electron에 기대지 않는
// 순수 함수만 두어(Node에서 바로 단위 테스트 가능) AI 없이 규칙으로만 동작한다.
//
// - 기준 컬럼 추천: 두 표에서 "값이 많이 겹치고, 추가 원본 쪽에서 값이 거의 겹치지 않는(=행을 하나로
//   가리키는) 컬럼 쌍"을 찾는다. 이름이 같거나 비슷하면 가산점.
// - 합치기: 첫 원본(A)의 행마다 기준값이 같은 추가 원본 행의 컬럼을 이어 붙인다(왼쪽 기준).
//   "양쪽 행 모두"를 고르면 결과에 들어가지 못한 추가 원본 행(짝 없음·기준값 빈 칸·겹치는 기준값)을
//   모두 뒤에 덧붙여 어느 쪽 행도 빠지지 않는다. 빠지는 행은 행 번호와 함께 보고한다.
// - 기준값은 기본으로 글자 그대로 비교한다("00123" ≠ "123"). 앞자리 0 무시는 사용자가 고르는 선택이고,
//   고르기 전에 analyzeKeyMatch로 연결이 몇 행 늘고 서로 다른 값이 몇 개 같아지는지 보여 준다.
// - 시트 나누기 추천: 값 종류가 적고(2~30) 대부분 채워진 컬럼, 이름이 부서·지점·구분 같은 컬럼.

const normalizeName = (s) => String(s).replace(/\s+/g, '').toLowerCase();
const isEmpty = (v) => v === '' || v === null || v === undefined;

function pad2(n) {
  return String(n).padStart(2, '0');
}

/**
 * 기준값 비교용 정규화. 날짜는 연-월-일(시각 포함)로, 나머지는 앞뒤 공백을 뺀 글자로 맞춘다
 * (엑셀 숫자 123과 글자 "123"은 같다). 앞자리 0은 기본으로 그대로 비교한다 — "00123"과 "123"은
 * 다른 코드일 수 있다. ignoreLeadingZeros를 켜면 숫자로만 된 값의 앞자리 0을 뺀다(한쪽은 엑셀
 * 숫자, 다른 쪽은 CSV 글자로 같은 코드가 들어온 경우).
 */
function normalizeKey(v, { ignoreLeadingZeros = false } = {}) {
  if (isEmpty(v)) return '';
  if (v instanceof Date) {
    const base = `${v.getFullYear()}-${pad2(v.getMonth() + 1)}-${pad2(v.getDate())}`;
    return v.getHours() || v.getMinutes() || v.getSeconds() ? `${base} ${pad2(v.getHours())}:${pad2(v.getMinutes())}:${pad2(v.getSeconds())}` : base;
  }
  const s = String(v).trim();
  if (ignoreLeadingZeros && /^\d+$/.test(s)) return s.replace(/^0+(?=\d)/, '');
  return s;
}

/**
 * 앞자리 0을 무시할지 고르기 전에 보여 줄 숫자. 정확히 비교할 때와 무시할 때 연결되는 원본(A) 행 수,
 * 그리고 무시하면 서로 다른 값이 같아지는 경우(예: "0123"과 "123"이 한쪽 파일 안에 함께 있음) 수.
 * 반환: { exactMatched, looseMatched, collisions, collisionExamples }
 */
function analyzeKeyMatch(main, keyMain, other, keyOther) {
  // 파일마다 "앞자리 0을 뺀 값 -> 원래 값들". 같은 파일 안에서 서로 다른 값이 같아지는 경우만 센다
  // (파일 사이에서 123 = "00123"이 되는 것은 의도한 연결이다)
  const profile = (table, key) => {
    const strict = new Set();
    const looseToStrict = new Map();
    const keys = table.rows.map((r) => {
      const s = normalizeKey(r[key]);
      if (s === '') return null;
      const l = normalizeKey(r[key], { ignoreLeadingZeros: true });
      strict.add(s);
      if (!looseToStrict.has(l)) looseToStrict.set(l, new Set());
      looseToStrict.get(l).add(s);
      return { s, l };
    });
    return { strict, looseToStrict, keys };
  };
  const pm = profile(main, keyMain);
  const po = profile(other, keyOther);
  let exactMatched = 0;
  let looseMatched = 0;
  pm.keys.forEach((k) => {
    if (!k) return;
    if (po.strict.has(k.s)) exactMatched += 1;
    if (po.looseToStrict.has(k.l)) looseMatched += 1;
  });
  let collisions = 0;
  const collisionExamples = [];
  [pm, po].forEach(({ looseToStrict }) => {
    looseToStrict.forEach((set) => {
      if (set.size < 2) return;
      collisions += 1;
      if (collisionExamples.length < 3) collisionExamples.push(Array.from(set).join(' · '));
    });
  });
  return { exactMatched, looseMatched, collisions, collisionExamples };
}

/** 이름 비슷한 정도(0~1): 같음 1 · 동의어 0.9 · 한쪽이 다른 쪽을 포함 0.6 */
function nameScore(a, b, synonymGroups = []) {
  const x = normalizeName(a);
  const y = normalizeName(b);
  if (!x || !y) return 0;
  if (x === y) return 1;
  if (synonymGroups.some((g) => g.some((t) => normalizeName(t) === x) && g.some((t) => normalizeName(t) === y))) return 0.9;
  if (x.length >= 2 && y.length >= 2 && (x.includes(y) || y.includes(x))) return 0.6;
  return 0;
}

const CODE_VALUE = /^(?=.*\d)[A-Za-z0-9._/-]+$/; // 사번·코드처럼 숫자가 들어간 영숫자 값
const KEY_NAME_HINT = /(번호|코드|사번|학번|아이디|id$|^id|no\.?$|key|번$)/i;

/** 컬럼마다 (앞 maxRows행 기준) 채워진 칸 수, 서로 다른 기준값 집합, 코드 모양 비율 */
function profileColumns(table, maxRows) {
  const n = Math.min(table.rows.length, maxRows);
  const profiles = {};
  table.headers.forEach((h) => {
    const distinct = new Set();
    const loose = new Set(); // 앞자리 0을 뺀 값(앞자리 0 무시를 켜야 이어지는 경우를 찾기 위해)
    let filled = 0;
    let codeLike = 0;
    for (let i = 0; i < n; i++) {
      const key = normalizeKey(table.rows[i][h]);
      if (key === '') continue;
      filled += 1;
      distinct.add(key);
      loose.add(normalizeKey(table.rows[i][h], { ignoreLeadingZeros: true }));
      if (CODE_VALUE.test(key)) codeLike += 1;
    }
    profiles[h] = { filled, distinct, loose, uniqueness: filled ? distinct.size / filled : 0, codeLike: filled ? codeLike / filled : 0 };
  });
  return profiles;
}

const KEY_MIN_OVERLAP = 0.3; // 값이 이만큼도 안 겹치면 후보로 보지 않는다
const KEY_MIN_UNIQUENESS = 0.7; // 추가 원본의 기준 컬럼은 값이 거의 겹치지 않아야 한다(한 값 = 한 행, 약간의 중복은 허용)

/**
 * 첫 원본(main)과 추가 원본(other)을 이을 기준 컬럼 쌍을 점수순으로 추천한다.
 * 점수 = 값 일치율 0.5 + 추가 원본 쪽 고유도 0.15 + 이름 비슷한 정도 0.15
 *      + 값이 코드 모양(사번·코드처럼 숫자가 든 영숫자) 0.15 + 이름이 번호·코드·ID 류 0.05
 * 값 일치율 = 두 컬럼의 서로 다른 값 중 겹치는 비율(작은 쪽 기준). 사람 이름처럼 우연히 겹칠 수
 * 있는 값보다 코드 모양 값을 기준으로 고르도록 코드 모양에 가산점을 준다.
 * 값은 글자 그대로 비교한다. 그대로는 안 겹치고 앞자리 0을 무시해야 겹치는 쌍은 needsIgnoreZeros로
 * 표시하고 점수를 조금 깎는다(앞자리 0 무시는 사용자가 숫자를 보고 직접 켠다).
 * 반환: [{ main, other, score, overlap, uniqueness, needsIgnoreZeros, reasons: [문자열] }] (최대 limit개)
 */
function suggestJoinKeys(main, other, { synonymGroups = [], maxRows = 5000, limit = 5 } = {}) {
  const pm = profileColumns(main, maxRows);
  const po = profileColumns(other, maxRows);
  const results = [];
  other.headers.forEach((ho) => {
    const o = po[ho];
    if (o.distinct.size < 2 || o.uniqueness < KEY_MIN_UNIQUENESS) return;
    main.headers.forEach((hm) => {
      const m = pm[hm];
      if (m.distinct.size < 2) return;
      const overlapOf = (x, y) => {
        const [small, large] = x.size <= y.size ? [x, y] : [y, x];
        let common = 0;
        small.forEach((k) => {
          if (large.has(k)) common += 1;
        });
        return common / small.size;
      };
      let overlap = overlapOf(m.distinct, o.distinct);
      let needsIgnoreZeros = false;
      if (overlap < KEY_MIN_OVERLAP) {
        const looseOverlap = overlapOf(m.loose, o.loose);
        if (looseOverlap < KEY_MIN_OVERLAP) return;
        overlap = looseOverlap;
        needsIgnoreZeros = true;
      }
      const nm = nameScore(hm, ho, synonymGroups);
      const codeLike = m.codeLike >= 0.9 && o.codeLike >= 0.9;
      const keyName = KEY_NAME_HINT.test(normalizeName(hm)) || KEY_NAME_HINT.test(normalizeName(ho));
      const score = overlap * 0.5 + o.uniqueness * 0.15 + nm * 0.15 + (codeLike ? 0.15 : 0) + (keyName ? 0.05 : 0) - (needsIgnoreZeros ? 0.05 : 0);
      const reasons = [needsIgnoreZeros ? `앞자리 0을 무시하면 값 일치 ${Math.round(overlap * 100)}%` : `값 일치 ${Math.round(overlap * 100)}%`];
      if (o.uniqueness >= 0.999) reasons.push('겹치는 값 없음');
      if (codeLike) reasons.push('코드 모양 값');
      if (nm === 1) reasons.push('이름 같음');
      else if (nm === 0.9) reasons.push('동의어');
      else if (nm > 0) reasons.push('이름 비슷');
      results.push({ main: hm, other: ho, score, overlap, uniqueness: o.uniqueness, needsIgnoreZeros, reasons });
    });
  });
  return results.sort((x, y) => y.score - x.score).slice(0, limit);
}

/**
 * 여러 원본을 하나의 표로 합친다.
 * main: { headers, labels, rows, rowNumbers, numberFormats }
 * others: [{ label, table, keyMain, keyOther, mode: 'left' | 'outer', dupPolicy: 'first' | 'all', ignoreLeadingZeros }]
 * - 추가 원본의 기준 컬럼은 결과에 다시 넣지 않는다(첫 원본의 기준 컬럼과 같은 값).
 * - 이름이 겹치는 컬럼은 "이름 (파일라벨)"로 구분한다.
 * - 추가 원본에서 기준값이 같은 행이 여럿일 때: dupPolicy 'first'면 첫 행만 잇고(나머지는 보고),
 *   'all'이면 짝이 되는 행마다 원본(A) 행을 하나씩 더 만든다(결과 행이 늘어난다).
 * - mode 'left': 결과 행은 원본(A) 행. 결과에 못 들어간 추가 원본 행은 이유별로 행 번호와 함께 보고한다.
 * - mode 'outer'(양쪽 행 모두): 결과에 못 들어간 추가 원본 행을 모두(짝 없음·기준값 빈 칸·겹치는 기준값)
 *   뒤에 덧붙인다 — 어느 쪽 행도 빠지지 않는다.
 * 반환: { table: { headers, labels, rows, rowNumbers, numberFormats, columnOrigin }, reports: [...] }
 *   rowNumbers는 첫 원본 행은 숫자, 추가 원본에서 덧붙인 행은 "라벨 행번호" 글자.
 */
function joinTables(main, others) {
  const headers = [...main.headers];
  const labels = [...(main.labels || main.headers)];
  const numberFormats = { ...(main.numberFormats || {}) };
  const columnOrigin = {};
  let rows = main.rows.map((r) => ({ ...r }));
  let rowNumbers = [...(main.rowNumbers || main.rows.map((_, i) => i + 1))];
  const used = new Set(headers);
  const reports = [];

  others.forEach(({ label, table, keyMain, keyOther, mode = 'left', dupPolicy = 'first', ignoreLeadingZeros = false }) => {
    const keyOptions = { ignoreLeadingZeros };
    const otherNumbers = table.rowNumbers || table.rows.map((_, i) => i + 1);

    // 새 컬럼 이름 정하기
    const renamed = {};
    table.headers.forEach((h, idx) => {
      if (h === keyOther) return;
      let name = used.has(h) ? `${h} (${label})` : h;
      for (let n = 1; used.has(name); n++) name = `${h} (${label})_${n}`;
      used.add(name);
      renamed[h] = name;
      headers.push(name);
      labels.push(table.labels ? table.labels[idx] : h);
      columnOrigin[name] = label;
      if (table.numberFormats && table.numberFormats[h]) numberFormats[name] = table.numberFormats[h];
    });
    const newColumns = Object.keys(renamed);
    const fill = (row, src) => {
      newColumns.forEach((h) => {
        row[renamed[h]] = src && src[h] !== undefined ? src[h] : '';
      });
      return row;
    };

    // 추가 원본을 기준값으로 찾을 수 있게: 기준값 -> [행 번호(0부터)]
    const index = new Map();
    const emptyKey = [];
    table.rows.forEach((r, j) => {
      const key = normalizeKey(r[keyOther], keyOptions);
      if (key === '') {
        emptyKey.push(j);
        return;
      }
      if (!index.has(key)) index.set(key, []);
      index.get(key).push(j);
    });
    let duplicateRows = 0;
    let duplicateKeys = 0;
    index.forEach((list) => {
      if (list.length > 1) {
        duplicateKeys += 1;
        duplicateRows += list.length - 1;
      }
    });

    const usedOther = new Set();
    const matchedKeys = new Set();
    const nextRows = [];
    const nextNumbers = [];
    const unmatched = [];
    let matchedRows = 0;
    let expandedRows = 0;
    rows.forEach((row, i) => {
      const key = normalizeKey(row[keyMain], keyOptions);
      const list = key === '' ? undefined : index.get(key);
      if (!list) {
        unmatched.push(rowNumbers[i]);
        nextRows.push(fill(row, null));
        nextNumbers.push(rowNumbers[i]);
        return;
      }
      matchedRows += 1;
      matchedKeys.add(key);
      (dupPolicy === 'all' ? list : list.slice(0, 1)).forEach((j, t) => {
        usedOther.add(j);
        nextRows.push(fill(t === 0 ? row : { ...row }, table.rows[j]));
        nextNumbers.push(rowNumbers[i]);
        if (t > 0) expandedRows += 1;
      });
    });
    rows = nextRows;
    rowNumbers = nextNumbers;

    // 결과에 못 들어간 추가 원본 행을 이유별로: 짝 없음 · 겹치는 기준값(첫 행만 쓸 때) · 기준값 빈 칸
    const otherOnly = [];
    const skippedDuplicates = [];
    index.forEach((list, key) => {
      if (!matchedKeys.has(key)) otherOnly.push(...list);
      else list.forEach((j) => {
        if (!usedOther.has(j)) skippedDuplicates.push(j);
      });
    });
    const leftOut = [...otherOnly, ...skippedDuplicates, ...emptyKey].sort((x, y) => x - y);
    if (mode === 'outer') {
      leftOut.forEach((j) => {
        const src = table.rows[j];
        const row = {};
        headers.forEach((h) => {
          row[h] = '';
        });
        row[keyMain] = src[keyOther] === undefined ? '' : src[keyOther];
        rows.push(fill(row, src));
        rowNumbers.push(`${label} ${otherNumbers[j]}`);
      });
    }
    const numbersOf = (list) => list.sort((x, y) => x - y).map((j) => otherNumbers[j]);

    reports.push({
      label,
      keyMain,
      keyOther,
      mode,
      dupPolicy,
      ignoreLeadingZeros,
      addedColumns: newColumns.map((h) => renamed[h]),
      matchedRows,
      unmatchedRows: unmatched.length,
      unmatchedRowNumbers: unmatched,
      expandedRows,
      duplicateKeys,
      duplicateRows,
      otherOnlyRows: otherOnly.length,
      otherOnlyRowNumbers: numbersOf(otherOnly),
      skippedDuplicateRows: skippedDuplicates.length,
      skippedDuplicateRowNumbers: numbersOf(skippedDuplicates),
      emptyKeyRows: emptyKey.length,
      emptyKeyRowNumbers: numbersOf(emptyKey),
      addedRows: mode === 'outer' ? leftOut.length : 0,
      excludedRows: mode === 'outer' ? 0 : leftOut.length,
    });
  });

  return { table: { headers, labels, rows, rowNumbers, numberFormats, columnOrigin }, reports };
}

/** 행 번호 목록을 짧게: "행 3, 7, 9 외 4개" */
function rowList(numbers, max = 5) {
  if (!numbers || numbers.length === 0) return '';
  const head = numbers.slice(0, max).join(', ');
  return `행 ${head}${numbers.length > max ? ` 외 ${numbers.length - max}개` : ''}`;
}

/**
 * 합치기 결과 한 건을 사람이 읽는 문장으로. 빠진 행·덧붙인 행은 행 번호와 함께.
 * 반환: { text, warn, details: [{ text, warn }] } — text는 한 줄 요약, details는 이유별 줄
 */
function describeJoinReportDetail(r) {
  const details = [];
  const outer = r.mode === 'outer';
  if (r.unmatchedRows) details.push({ text: `짝이 없는 원본(A) 행 ${r.unmatchedRows}개 — 추가 컬럼은 빈칸 (${rowList(r.unmatchedRowNumbers)})`, warn: true });
  if (r.otherOnlyRows) {
    details.push({
      text: outer ? `${r.label}에만 있는 행 ${r.otherOnlyRows}개 뒤에 덧붙임 (${rowList(r.otherOnlyRowNumbers)})` : `${r.label}에만 있는 행 ${r.otherOnlyRows}개는 넣지 않음 (${rowList(r.otherOnlyRowNumbers)})`,
      warn: !outer,
    });
  }
  if (r.duplicateRows) {
    if (r.dupPolicy === 'all') {
      details.push({ text: `${r.label}에서 기준값이 겹치는 값 ${r.duplicateKeys}종류 — 짝마다 행을 만들어 결과 행 ${r.expandedRows}개 늘어남`, warn: true });
    } else if (r.skippedDuplicateRows) {
      details.push({
        text: outer
          ? `${r.label}에서 기준값이 겹쳐 첫 행만 이은 값 ${r.duplicateKeys}종류 — 나머지 ${r.skippedDuplicateRows}행은 따로 덧붙임 (${rowList(r.skippedDuplicateRowNumbers)})`
          : `${r.label}에서 기준값이 겹쳐 첫 행만 이음 — 나머지 ${r.skippedDuplicateRows}행은 넣지 않음 (${rowList(r.skippedDuplicateRowNumbers)})`,
        warn: true,
      });
    }
  }
  if (r.emptyKeyRows) {
    details.push({
      text: outer ? `${r.label}에서 기준값이 빈 행 ${r.emptyKeyRows}개 뒤에 덧붙임 (${rowList(r.emptyKeyRowNumbers)})` : `${r.label}에서 기준값이 빈 행 ${r.emptyKeyRows}개는 넣지 않음 (${rowList(r.emptyKeyRowNumbers)})`,
      warn: !outer,
    });
  }
  const summary = [`${r.label}: 기준 ${r.keyMain} = ${r.keyOther}${r.ignoreLeadingZeros ? '(앞자리 0 무시)' : ''}`, `연결된 행 ${r.matchedRows}개`];
  if (r.excludedRows) summary.push(`${r.label} 행 ${r.excludedRows}개 빠짐`);
  if (r.addedRows) summary.push(`${r.label} 행 ${r.addedRows}개 덧붙임`);
  return { text: summary.join(' · '), warn: details.some((d) => d.warn), details };
}

/** 합치기 결과 한 건을 한 문단으로(로그용) */
function describeJoinReport(r) {
  const d = describeJoinReportDetail(r);
  return [d.text, ...d.details.map((x) => x.text)].join(' · ');
}

const SPLIT_NAME_HINT = /(부서|지점|팀|소속|구분|분류|지역|월|유형|종류|상태|직급|직위|센터|본부|그룹|학년|반|category|type|dept|team|region|branch|group|status)/i;
const SPLIT_MAX_VALUES = 30;

/**
 * 결과를 시트 여러 개로 나눌 때 쓸 만한 타겟 컬럼을 추천한다.
 * valueOf(i, b)는 저장될 값의 비교용 글자. 값 종류가 2~30개이고 80% 이상 채워졌으며, 값이 행마다
 * 다르지 않은(같은 값이 반복되고 종류가 행 수의 60% 이하인) 컬럼만 후보로 본다.
 * 반환: [{ column, values, score, reasons }] 점수순
 */
function suggestSplitColumns(columns, rowCount, valueOf) {
  if (rowCount < 2) return [];
  const results = [];
  columns.forEach((b) => {
    const distinct = new Set();
    let filled = 0;
    for (let i = 0; i < rowCount; i++) {
      const v = valueOf(i, b);
      if (v === '') continue;
      filled += 1;
      distinct.add(v);
      if (distinct.size > SPLIT_MAX_VALUES) return;
    }
    const filledRatio = filled / rowCount;
    // 값이 행마다 다르면(이름·번호 같은 컬럼) 나눌 의미가 없다: 같은 값이 한 번은 반복되고, 종류가 행 수의 60% 이하
    if (distinct.size < 2 || filledRatio < 0.8 || distinct.size >= filled || distinct.size > Math.max(2, rowCount * 0.6)) return;
    const hint = SPLIT_NAME_HINT.test(b);
    const score = (hint ? 0.5 : 0) + (1 - distinct.size / SPLIT_MAX_VALUES) * 0.3 + filledRatio * 0.2;
    const reasons = [`값 ${distinct.size}종류`];
    if (hint) reasons.push('나누기 좋은 이름');
    if (filledRatio < 1) reasons.push(`빈칸 ${Math.round((1 - filledRatio) * 100)}%`);
    results.push({ column: b, values: distinct.size, score, reasons });
  });
  return results.sort((x, y) => y.score - x.score);
}

const SHEET_NAME_MAX = 31;

/** 엑셀 시트 이름 규칙에 맞게: 금지 문자 : \ / ? * [ ] 를 _로, 31자 이내, 겹치면 " (2)" */
function sanitizeSheetName(name, used) {
  let base = String(name).replace(/[:\\/?*[\]]/g, '_').replace(/^'+|'+$/g, '').trim();
  if (!base) base = '(빈 값)';
  if (base.toLowerCase() === 'history') base = `${base} (시트)`;
  base = base.slice(0, SHEET_NAME_MAX);
  let candidate = base;
  for (let n = 2; used.has(candidate.toLowerCase()); n++) {
    const suffix = ` (${n})`;
    candidate = base.slice(0, SHEET_NAME_MAX - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

const BLANK_GROUP_LABEL = '(빈 값)';

/**
 * 행 번호들을 keyOf(i) 값별로 묶는다(처음 나온 순서).
 * 빈 칸 묶음의 key는 ''이고 화면·시트 이름에는 label "(빈 값)"을 쓴다 — 실제로 "(빈 값)"이라고
 * 적힌 칸과 섞이지 않는다(그 경우 시트 이름은 sanitizeSheetName이 "(빈 값) (2)"처럼 구분한다).
 * 반환: [{ key, label, isBlank, indices }]
 */
function groupRowsBy(rowCount, keyOf) {
  const groups = new Map();
  for (let i = 0; i < rowCount; i++) {
    const v = keyOf(i);
    const key = v === undefined || v === null ? '' : String(v);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(i);
  }
  return Array.from(groups, ([key, indices]) => ({ key, label: key === '' ? BLANK_GROUP_LABEL : key, isBlank: key === '', indices }));
}

module.exports = {
  normalizeKey,
  analyzeKeyMatch,
  nameScore,
  suggestJoinKeys,
  joinTables,
  describeJoinReport,
  describeJoinReportDetail,
  suggestSplitColumns,
  sanitizeSheetName,
  groupRowsBy,
};
