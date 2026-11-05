// print-plan.js
// 서류 출력 화면의 판단 규칙: 기관·구분에 맞는 서식, 지원 종류별 묶음, 고른 서식에 필요한 입력과 비어 있는 칸,
// 출력 전 확인, 채운 파일 이름. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.

const { INPUT_BY_KEY, INPUT_GROUPS, inputsFor, missingInputs, checkRrn } = require('./merge-fields.js');
const { KINDS } = require('./form-library.js');

/** 그 기관의 서식 중 고른 구분(세대주/세대원)과 공통 서식. 지원 종류 순 → ★ 먼저 → 이름 순 */
function entriesFor(entries, agency, applicant) {
  return entries
    .filter((e) => e.agency === agency && (e.applicant === applicant || e.applicant === '공통'))
    .sort((a, b) => KINDS.indexOf(a.kind) - KINDS.indexOf(b.kind) || b.favorite - a.favorite || a.title.localeCompare(b.title, 'ko'));
}

/** 지원 종류별 묶음(목록 순서 그대로): [{ kind, items }] */
function groupByKind(list) {
  const groups = [];
  list.forEach((e) => {
    const g = groups.find((x) => x.kind === e.kind);
    if (g) g.items.push(e);
    else groups.push({ kind: e.kind, items: [e] });
  });
  return groups;
}

const mappingType = (entry, name) => (entry.mapping[name] || { type: 'auto' }).type;

/** 사전 값으로 채울 필드 이름들(매핑이 auto) */
const autoFieldNames = (list) => list.flatMap((e) => e.fields.filter((f) => mappingType(e, f.name) === 'auto').map((f) => f.name));

/** 출력할 때 물어볼 필드 이름들(매핑이 ask, 겹치면 한 번) */
const askFieldNames = (list) => [...new Set(list.flatMap((e) => e.fields.filter((f) => (e.mapping[f.name] || {}).type === 'ask').map((f) => f.name)))];

/** 고른 서식들에 필요한 입력 key(INPUTS 순서) */
const neededInputs = (list) => inputsFor(autoFieldNames(list)).keys;

/** 서식 하나에 비어 있는 칸의 이름표(입력 항목 이름, 출력 때 입력하는 필드 이름). input에는 설정의 담당자도 넣어 준다 */
function missingFor(entry, input, asks = {}) {
  const miss = missingInputs(neededInputs([entry]), input).map((k) => INPUT_BY_KEY.get(k).label);
  askFieldNames([entry]).forEach((n) => {
    if (!String(asks[n] || '').trim()) miss.push(n);
  });
  return miss;
}

/** 여러 서식에서 비어 있는 칸(겹치면 한 번) */
const missingForAll = (list, input, asks) => [...new Set(list.flatMap((e) => missingFor(e, input, asks)))];

/** 설정에서 넣어야 하는데 비어 있는 담당자 칸(이름표) */
const missingStaff = (list, settings) =>
  neededInputs(list)
    .filter((k) => INPUT_GROUPS.staff.includes(k) && !String(settings[k] || '').trim())
    .map((k) => INPUT_BY_KEY.get(k).label);

/**
 * 출력 전 확인. 반환: { block: '막는 이유' | '', focus: 입력 key | null, missing: [비운 채로 나갈 칸] }
 * - 필요한 필수 칸(이름·주민등록번호)이 비었거나 주민등록번호 모양이 틀리면 막는다
 * - 그 밖의 빈칸은 알리고 계속할지 묻는다(missing)
 */
function checkBeforeOutput(list, input, asks) {
  const need = neededInputs(list).filter((k) => INPUT_BY_KEY.get(k).required && !String(input[k] || '').trim());
  if (need.length > 0) return { block: `${need.map((k) => INPUT_BY_KEY.get(k).label).join(', ')}을(를) 넣어 주세요.`, focus: need[0], missing: [] };
  const rrnProblem = checkRrn(input.rrn || '');
  if (rrnProblem) return { block: `주민등록번호를 확인하세요: ${rrnProblem}`, focus: 'rrn', missing: [] };
  return { block: '', focus: null, missing: missingForAll(list, input, asks) };
}

/** 출력 줄 요약: { title, detail } */
function summaryOf(list, input, asks, copies) {
  if (list.length === 0) return { title: '출력할 서류를 고르세요', detail: '' };
  const docCount = list.reduce((n, e) => n + (e.documents.length || 1), 0);
  const rrnProblem = checkRrn(input.rrn || '');
  const miss = missingForAll(list, input, asks);
  return {
    title: `서식 ${list.length}개, 서류 약 ${docCount}종, ${copies}부씩`,
    detail: rrnProblem ? `주민등록번호: ${rrnProblem}` : miss.length > 0 ? `빈칸으로 나갈 칸: ${miss.join(', ')}` : '모든 칸이 채워집니다',
  };
}

/** 채운 서식을 저장할 파일 이름: "기관_서식 이름(구분).hwpx"(파일 이름에 못 쓰는 글자는 _) */
const filledFileName = (entry) => `${entry.agency}_${entry.title}${entry.applicant !== '공통' ? `(${entry.applicant})` : ''}.hwpx`.replace(/[\\/:*?"<>|]/g, '_');

module.exports = { entriesFor, groupByKind, autoFieldNames, askFieldNames, neededInputs, missingFor, missingForAll, missingStaff, checkBeforeOutput, summaryOf, filledFileName };
