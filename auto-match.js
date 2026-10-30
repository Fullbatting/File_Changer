// auto-match.js
// 컬럼 이름 자동 연결(타겟 B 컬럼 ← 원본 A 컬럼). 화면에 기대지 않는 순수 함수라 node로 단위 테스트한다
// (tests/unit/auto-match.test.js). AI 없이 정해진 규칙으로만 판단하므로 같은 입력은 언제나 같은 결과다.

const FUZZY_MAX_DISTANCE = 2; // 절대 편집거리 상한
const FUZZY_MAX_RATIO = 0.34; // 상대 편집거리 상한(짧은 이름에서 과매칭 방지)
const JAMO_FUZZY_MAX_DISTANCE = 2; // 절대 편집거리 상한(자모 단위)
const JAMO_FUZZY_MAX_RATIO = 0.3; // 음절 단위(0.34)보다 더 엄격하게 잡아 과매칭을 방지

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

/** 문자열 비교용 정규화(공백 제거 + 소문자 변환). 자동 매핑/동의어 사전에서 공통으로 사용 */
function normalizeFieldName(s) {
  return String(s).replace(/\s+/g, '').toLowerCase();
}

/* -------------------------------------------------------------------------
 * 한글 자모 분해 (음절 단위 퍼지 매칭이 놓치는 한글 오타 보완)
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

/** 한글 음절을 초성으로 바꾼다("연락처" → "ㅇㄹㅊ"). 한글이 아닌 문자는 그대로 둔다 */
function getChoseong(str) {
  let result = '';
  for (const ch of String(str)) {
    const code = ch.codePointAt(0);
    if (code >= HANGUL_SYLLABLE_BASE && code <= HANGUL_SYLLABLE_LAST) {
      result += HANGUL_CHOSEONG[Math.floor((code - HANGUL_SYLLABLE_BASE) / (21 * 28))];
    } else {
      result += ch;
    }
  }
  return result;
}

/* ---------- 단계별 후보 ---------- */

/** 정규화 이름이 같은 A 후보. 원래 글자까지 똑같은 것이 하나면 그것을 고른다 */
function exactCandidates(bCol, available) {
  const bNorm = normalizeFieldName(bCol);
  const same = available.filter((a) => normalizeFieldName(a) === bNorm);
  if (same.length <= 1) return same;
  const identical = same.filter((a) => a === bCol);
  return identical.length === 1 ? identical : same;
}

/** 동의어 사전에서 bCol과 같은 그룹인 A 후보 */
function synonymCandidates(bCol, available, synonymGroups) {
  const bNorm = normalizeFieldName(bCol);
  const group = synonymGroups.find((g) => g.some((term) => normalizeFieldName(term) === bNorm));
  if (!group) return [];
  const groupNormSet = new Set(group.map(normalizeFieldName));
  return available.filter((a) => groupNormSet.has(normalizeFieldName(a)) && normalizeFieldName(a) !== bNorm);
}

/** 편집거리 기준을 넘지 않는 A 후보를 가까운 순으로: [{ a, dist }] */
function fuzzyCandidates(bCol, available, jamo) {
  const conv = jamo ? (s) => decomposeHangulToJamo(normalizeFieldName(s)) : normalizeFieldName;
  const maxDist = jamo ? JAMO_FUZZY_MAX_DISTANCE : FUZZY_MAX_DISTANCE;
  const maxRatio = jamo ? JAMO_FUZZY_MAX_RATIO : FUZZY_MAX_RATIO;
  const bKey = conv(bCol);
  return available
    .map((a) => {
      const aKey = conv(a);
      const dist = levenshteinDistance(aKey, bKey);
      return { a, dist, ratio: dist / Math.max(aKey.length, bKey.length, 1) };
    })
    .filter((c) => c.dist > 0 && c.dist <= maxDist && c.ratio <= maxRatio)
    .sort((x, y) => x.dist - y.dist);
}

/** 가장 가까운 후보가 하나뿐이면 그 후보, 동률이면 null */
function uniqueBest(list) {
  if (list.length === 0) return null;
  if (list.length === 1 || list[0].dist < list[1].dist) return list[0];
  return null;
}

const tiedOf = (list) => (list.length > 1 ? list.filter((c) => c.dist === list[0].dist).map((c) => c.a) : []);

/**
 * B 컬럼 하나만 볼 때의 자동 연결(메일머지 변수처럼 한 번에 하나씩 정할 때).
 * 1순위 같은 이름(정규화 후) → 2순위 동의어 사전(후보가 하나일 때만) → 3순위 비슷한 이름(음절)
 * → 4순위 비슷한 이름(자모). 어느 단계든 후보가 동률이면 연결하지 않고 candidates로 돌려준다 —
 * 잘못 연결되어 데이터가 다른 컬럼에 들어가는 것이 연결이 안 되는 것보다 더 나쁘다.
 * usedHeadersA에 든 컬럼은 이미 다른 B에 연결되었으므로 후보에서 뺀다.
 * 여러 B를 한꺼번에 정할 때는 순서에 따라 결과가 달라지지 않도록 autoMatchAll을 쓴다.
 */
function findAutoMatch(bCol, headersA, usedHeadersA, synonymGroups = []) {
  const available = headersA.filter((a) => !usedHeadersA.has(a));

  const exact = exactCandidates(bCol, available);
  if (exact.length === 1) return { match: exact[0], reason: 'exact', candidates: [] };
  if (exact.length > 1) return { match: null, reason: null, candidates: exact };

  const syn = synonymCandidates(bCol, available, synonymGroups);
  if (syn.length === 1) return { match: syn[0], reason: 'synonym', candidates: [] };
  if (syn.length > 1) return { match: null, reason: null, candidates: syn };

  const syllable = fuzzyCandidates(bCol, available, false);
  const best = uniqueBest(syllable);
  if (best) return { match: best.a, reason: 'fuzzy', candidates: [] };
  const jamo = fuzzyCandidates(bCol, available, true);
  const bestJamo = uniqueBest(jamo);
  if (bestJamo) return { match: bestJamo.a, reason: 'fuzzy', candidates: [] };

  const syllableTie = tiedOf(syllable);
  return { match: null, reason: null, candidates: syllableTie.length > 0 ? syllableTie : tiedOf(jamo) };
}

/**
 * 타겟 컬럼 전체를 한꺼번에 자동 연결한다. 타겟 컬럼 순서와 상관없이 같은 결과가 나오도록
 * 단계마다 "모든 B"를 먼저 본 뒤 확정한다:
 * 1) 같은 이름을 모두 먼저 확정. 그래서 타겟이 [성명, 이름]이고 원본에 "이름"만 있으면
 *    "이름"끼리 연결되고 "성명"은 동의어로 "이름"을 가져가지 못한다.
 *    정규화 이름이 같은 A가 여럿이면(원래 글자까지 같은 것이 하나가 아니면) 모호 — 연결하지 않는다.
 *    같은 A를 원하는 B가 여럿이어도 원래 글자까지 같은 B가 하나가 아니면 모호.
 * 2) 남은 B에 동의어 → 3) 비슷한 이름(음절) → 4) 비슷한 이름(자모)를 차례로.
 *    각 단계는 "B의 가장 가까운 후보가 하나뿐 + 그 A를 원하는 B 중 이 B가 가장 가까움(동률 아님)"인
 *    쌍만 확정하고, 확정으로 후보가 줄어 새로 하나뿐이 된 B를 위해 바뀌지 않을 때까지 되풀이한다.
 * 반환: { [b]: { match, reason, candidates, note } } — candidates는 모호해서 연결하지 않은 후보,
 * note는 왜 연결하지 않았는지(화면에서 후보와 함께 보여 준다)
 */
function autoMatchAll(headersB, headersA, synonymGroups = []) {
  const result = {};
  headersB.forEach((b) => {
    result[b] = { match: null, reason: null, candidates: [], note: '' };
  });
  const setAmbiguous = (b, candidates, note) => {
    if (result[b].candidates.length > 0) return; // 앞 단계의 이유를 남긴다
    result[b].candidates = candidates;
    result[b].note = note;
  };
  const usedA = new Set();
  // 한 단계에서 모호했던 B는 다음 단계로 넘기지 않는다(같은 이름이 둘이면 동의어로 엉뚱한 컬럼을 잡지 않도록)
  const blocked = new Set();
  const pending = () => headersB.filter((b) => !result[b].match && !blocked.has(b));
  const blockAmbiguous = () => pending().forEach((b) => {
    if (result[b].candidates.some((a) => !usedA.has(a))) blocked.add(b);
  });
  const available = () => headersA.filter((a) => !usedA.has(a));

  // 1) 같은 이름
  const wantsA = new Map(); // a -> [b]
  headersB.forEach((b) => {
    const list = exactCandidates(b, headersA);
    if (list.length === 1) {
      if (!wantsA.has(list[0])) wantsA.set(list[0], []);
      wantsA.get(list[0]).push(b);
    } else if (list.length > 1) {
      setAmbiguous(b, list, `이름이 같은 원본 컬럼 ${list.length}개`);
    }
  });
  wantsA.forEach((bs, a) => {
    const winner = bs.length === 1 ? bs[0] : bs.filter((b) => b === a).length === 1 ? a : null;
    if (winner) {
      result[winner] = { match: a, reason: 'exact', candidates: [], note: '' };
      usedA.add(a);
    } else {
      bs.forEach((b) => setAmbiguous(b, [a], `타겟 컬럼 ${bs.filter((x) => x !== b).map((x) => `"${x}"`).join(', ')}도 같은 이름`));
    }
  });

  blockAmbiguous();

  // 2~4) 남은 B에 단계별로. 후보 목록은 [{ a, dist }] (동의어는 dist 0으로 같게 본다)
  const phases = [
    { reason: 'synonym', kind: '동의어', list: (b, avail) => synonymCandidates(b, avail, synonymGroups).map((a) => ({ a, dist: 0 })) },
    { reason: 'fuzzy', kind: '비슷한 이름', list: (b, avail) => fuzzyCandidates(b, avail, false) },
    { reason: 'fuzzy', kind: '비슷한 이름', list: (b, avail) => fuzzyCandidates(b, avail, true) },
  ];
  phases.forEach(({ reason, kind, list }) => {
    for (;;) {
      const avail = available();
      const proposals = new Map(); // a -> [{ b, dist }]
      pending().forEach((b) => {
        const cands = list(b, avail);
        const best = uniqueBest(cands);
        if (best) {
          if (!proposals.has(best.a)) proposals.set(best.a, []);
          proposals.get(best.a).push({ b, dist: best.dist });
        } else if (cands.length > 1) {
          const tied = tiedOf(cands);
          setAmbiguous(b, tied, `${kind} 후보 ${tied.length}개가 같은 정도`);
        }
      });
      let assigned = 0;
      proposals.forEach((bs, a) => {
        const sorted = bs.slice().sort((x, y) => x.dist - y.dist);
        if (sorted.length > 1 && sorted[0].dist === sorted[1].dist) {
          const tiedBs = sorted.filter((x) => x.dist === sorted[0].dist).map((x) => x.b);
          tiedBs.forEach((b) => setAmbiguous(b, [a], `타겟 컬럼 ${tiedBs.filter((x) => x !== b).map((x) => `"${x}"`).join(', ')}도 같은 정도로 이 ${kind}`));
          return;
        }
        result[sorted[0].b] = { match: a, reason, candidates: [], note: '' };
        usedA.add(a);
        assigned += 1;
      });
      if (assigned === 0) break;
    }
    blockAmbiguous();
  });

  // 후보 목록에서 이미 다른 B에 연결된 A는 뺀다(고를 수는 있지만 추천으로는 의미가 없다)
  headersB.filter((b) => !result[b].match).forEach((b) => {
    result[b].candidates = result[b].candidates.filter((a) => !usedA.has(a));
    if (result[b].candidates.length === 0) result[b].note = '';
  });
  return result;
}

module.exports = {
  FUZZY_MAX_DISTANCE,
  FUZZY_MAX_RATIO,
  levenshteinDistance,
  normalizeFieldName,
  decomposeHangulToJamo,
  getChoseong,
  findAutoMatch,
  autoMatchAll,
};
