// merge-fields.js
// 한글 메일머지 필드({{이름}}·{{1}}~{{13}}·{{월1}} …) 사전과, 화면 입력 → 필드 값 규칙.
// 필드 이름과 값 모양은 지금까지 쓰던 메일머지 자료(엑셀 열·서식 필드)를 그대로 따른다:
//   생년월일 "900101"(6자리) · 주민등록번호 "900101-1234567" · {{1}}~{{13}} 주민등록번호 한 자리씩 ·
//   {{월1}}{{월2}}월 {{일1}}{{일2}}일 작성일 두 자리씩 · 진료기간 "26.01.02~" · 진단일 "26.01.01"
// 환자 인적사항은 저장하지 않는다(이 모듈은 값을 계산만 한다). 화면(DOM)에 기대지 않아 node로 단위 테스트한다.

/**
 * 화면 입력 항목. group: patient(환자 — 저장 안 함) · care(진료 — 저장 안 함) · staff(담당자 — PC 설정에 저장)
 * type: text | rrn | date | tel
 */
const INPUTS = [
  { key: 'name', group: 'patient', label: '이름', type: 'text', required: true },
  { key: 'rrn', group: 'patient', label: '주민등록번호', type: 'rrn', required: true, placeholder: '000000-0000000' },
  { key: 'addr', group: 'patient', label: '주소', type: 'text' },
  { key: 'phone', group: 'patient', label: '연락처', type: 'tel', placeholder: '010-0000-0000' },
  { key: 'writtenDate', group: 'patient', label: '작성일', type: 'date', hint: '비워 두면 오늘' },
  { key: 'careStart', group: 'care', label: '진료 시작일', type: 'date', hint: '진료기간 "26.01.02~" 모양으로 들어갑니다' },
  { key: 'careEnd', group: 'care', label: '진료 종료일', type: 'date', hint: '비워 두면 "~"로 끝남' },
  { key: 'disease', group: 'care', label: '주질병(상병명)', type: 'text' },
  { key: 'diseaseCode', group: 'care', label: '질병분류기호', type: 'text', placeholder: 'C00.0' },
  { key: 'diagnosisDate', group: 'care', label: '진단일', type: 'date' },
  { key: 'staffName', group: 'staff', label: '담당자', type: 'text' },
  { key: 'staffTitle', group: 'staff', label: '직위', type: 'text' },
  { key: 'staffPhone', group: 'staff', label: '담당자 번호', type: 'tel' },
];
const INPUT_BY_KEY = new Map(INPUTS.map((i) => [i.key, i]));
/** 묶음별 입력 key: patient(인적사항 칸) · care(진료 정보 칸) · staff(설정의 담당자) */
const keysOf = (group) => INPUTS.filter((i) => i.group === group).map((i) => i.key);
const INPUT_GROUPS = { patient: keysOf('patient'), care: keysOf('care'), staff: keysOf('staff') };
/** 비어 있어도 되는 입력(작성일은 오늘, 진료 종료일은 "~"로 끝남) */
const OPTIONAL_KEYS = ['writtenDate', 'careEnd'];

/** 시험 저장·한글 연결 점검에 쓰는 가짜 예시 값(실제 인적사항이 아님) */
const SAMPLE_INPUT = Object.freeze({
  name: '홍길동',
  rrn: '900101-1234567',
  addr: '[예시 주소] ○○시 ○○로 00',
  phone: '010-0000-0000',
  careStart: '2026-01-02',
  disease: '[예시 상병명]',
  diseaseCode: 'C00.0',
  diagnosisDate: '2026-01-01',
  staffName: '[담당자]',
  staffTitle: '[직위]',
  staffPhone: '0000',
});

/**
 * 필드 이름 → 값 만드는 법과 그 값에 필요한 입력. 별칭(names)은 다른 서식에서 같은 뜻으로 쓰는 이름.
 * value(ctx): ctx = { input, rrn: { digits, birth6, ok }, written: { y, m, d } }
 */
const FIELDS = [
  { id: 'name', names: ['이름', '성명', '대상자성명', '환자명', '성함'], needs: ['name'], value: (c) => c.input.name },
  { id: 'rrn', names: ['주민등록번호', '주민번호'], needs: ['rrn'], value: (c) => (c.rrn.ok ? `${c.rrn.digits.slice(0, 6)}-${c.rrn.digits.slice(6)}` : c.input.rrn) },
  { id: 'birth6', names: ['생년월일'], needs: ['rrn'], value: (c) => c.rrn.birth6 },
  ...Array.from({ length: 13 }, (_, i) => ({ id: `rrn${i + 1}`, names: [String(i + 1)], needs: ['rrn'], value: (c) => c.rrn.digits[i] || '' })),
  { id: 'addr', names: ['주소', '현주소', '주소지'], needs: ['addr'], value: (c) => c.input.addr },
  { id: 'phone', names: ['연락처', '전화번호', '휴대폰', '휴대폰번호', '휴대전화', '핸드폰'], needs: ['phone'], value: (c) => c.input.phone },
  { id: 'month1', names: ['월1'], needs: ['writtenDate'], value: (c) => pad2(c.written.m)[0] },
  { id: 'month2', names: ['월2'], needs: ['writtenDate'], value: (c) => pad2(c.written.m)[1] },
  { id: 'day1', names: ['일1'], needs: ['writtenDate'], value: (c) => pad2(c.written.d)[0] },
  { id: 'day2', names: ['일2'], needs: ['writtenDate'], value: (c) => pad2(c.written.d)[1] },
  { id: 'carePeriod', names: ['진료기간'], needs: ['careStart', 'careEnd'], value: (c) => carePeriod(c.input.careStart, c.input.careEnd) },
  { id: 'disease', names: ['주질병', '상병명', '진단명', '병명'], needs: ['disease'], value: (c) => c.input.disease },
  { id: 'diseaseCode', names: ['질병분류기호', '상병코드', '질병코드'], needs: ['diseaseCode'], value: (c) => c.input.diseaseCode },
  { id: 'diagnosisDate', names: ['진단일', '최초진단일', '진단일자'], needs: ['diagnosisDate'], value: (c) => shortDate(c.input.diagnosisDate) },
  { id: 'staffName', names: ['담당자', '담당자명'], needs: ['staffName'], value: (c) => c.input.staffName },
  { id: 'staffTitle', names: ['직위', '직책'], needs: ['staffTitle'], value: (c) => c.input.staffTitle },
  { id: 'staffPhone', names: ['담당자 번호', '담당자번호', '담당자 연락처', '담당자연락처'], needs: ['staffPhone'], value: (c) => c.input.staffPhone },
];

const norm = (s) => String(s || '').replace(/\s+/g, '').toLowerCase();
const FIELD_BY_NAME = new Map();
FIELDS.forEach((f) => f.names.forEach((n) => FIELD_BY_NAME.set(norm(n), f)));
const FIELD_BY_ID = new Map(FIELDS.map((f) => [f.id, f]));

const pad2 = (n) => String(n || '').padStart(2, '0');

/** 날짜 입력("2026-01-02" · "2026.1.2" · "26.01.02") → { y, m, d } 또는 null */
function parseDate(s) {
  const t = String(s || '').trim();
  let m = /^(\d{4})[-./\s]+(\d{1,2})[-./\s]+(\d{1,2})\.?$/.exec(t);
  if (m) return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
  m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})\.?$/.exec(t);
  if (m) return { y: 2000 + Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
  m = /^(\d{4})(\d{2})(\d{2})$/.exec(t);
  if (m) return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
  return null;
}

/** "26.01.02" 모양(날짜로 읽을 수 없으면 적은 그대로) */
function shortDate(s) {
  const d = parseDate(s);
  return d ? `${pad2(d.y % 100)}.${pad2(d.m)}.${pad2(d.d)}` : String(s || '').trim();
}

/** 진료기간 "26.01.02~" 또는 "26.01.02~26.01.31" */
function carePeriod(start, end) {
  if (!String(start || '').trim()) return '';
  return `${shortDate(start)}~${String(end || '').trim() ? shortDate(end) : ''}`;
}

/** 주민등록번호 → { digits(13자리 글자), birth6, ok(형식이 맞는지) } */
function rrnParts(rrn) {
  const digits = String(rrn || '').replace(/\D/g, '');
  return { digits, birth6: digits.slice(0, 6), ok: digits.length === 13 };
}

/** 주민등록번호 확인: 13자리 · 생년월일이 날짜인지. 반환: 문제 설명 또는 '' */
function checkRrn(rrn) {
  const t = String(rrn || '').trim();
  if (!t) return '';
  const { digits } = rrnParts(t);
  if (digits.length !== 13) return '주민등록번호는 숫자 13자리입니다';
  const mo = Number(digits.slice(2, 4));
  const d = Number(digits.slice(4, 6));
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '앞 6자리가 생년월일이 아닙니다';
  if (!/[1-8]/.test(digits[6]) && !/[90]/.test(digits[6])) return '뒤 첫 자리를 확인하세요';
  return '';
}

/** 주민등록번호로 성별·만 나이(화면 안내용) */
function rrnSummary(rrn, today = new Date()) {
  if (!rrn || checkRrn(rrn)) return '';
  const { digits } = rrnParts(rrn);
  const g = Number(digits[6]);
  const century = [1, 2, 5, 6].includes(g) ? 1900 : [3, 4, 7, 8].includes(g) ? 2000 : 1800;
  const y = century + Number(digits.slice(0, 2));
  const mo = Number(digits.slice(2, 4));
  const d = Number(digits.slice(4, 6));
  let age = today.getFullYear() - y;
  if (today.getMonth() + 1 < mo || (today.getMonth() + 1 === mo && today.getDate() < d)) age -= 1;
  return `생년월일 ${digits.slice(0, 6)}, ${g % 2 === 1 ? '남' : '여'}, 만 ${age}세`;
}

/** 필드 이름 → 사전 항목(없으면 null) */
const fieldOf = (name) => FIELD_BY_NAME.get(norm(name)) || null;

/** 사전에 없는 필드 이름들(fields: [{ name }] 또는 이름 목록) */
const unknownFields = (fields) => fields.map((f) => (typeof f === 'string' ? f : f.name)).filter((n) => !fieldOf(n));

/** 사전 필드가 어느 입력에서 오는지("이름", "진료 시작일·진료 종료일") */
const sourceLabel = (field) => field.needs.map((k) => INPUT_BY_KEY.get(k).label).join('·');

/**
 * 입력 → { 필드 이름: 값 } 함수. names: 서식에 든 필드 이름들(사전에 없는 이름은 extra[이름] 또는 빈칸)
 * input: 화면 입력(INPUTS의 key) · today: 작성일을 비웠을 때 쓸 날짜
 */
function fieldValues(names, input = {}, { today = new Date(), extra = {} } = {}) {
  const written = parseDate(input.writtenDate) || { y: today.getFullYear(), m: today.getMonth() + 1, d: today.getDate() };
  const ctx = { input: Object.fromEntries(Object.entries(input).map(([k, v]) => [k, String(v == null ? '' : v).trim()])), rrn: rrnParts(input.rrn), written };
  const out = {};
  names.forEach((name) => {
    const f = fieldOf(name);
    out[name] = f ? String(f.value(ctx) || '') : String(extra[name] == null ? '' : extra[name]);
  });
  out.__written = written;
  return out;
}

/**
 * 서식들에 든 필드 이름 → 물어야 할 입력 항목 key(INPUTS 순서)와, 사전에 없는 필드 이름들
 */
function inputsFor(names) {
  const keys = new Set();
  const unknown = [];
  names.forEach((name) => {
    const f = fieldOf(name);
    if (f) f.needs.forEach((k) => keys.add(k));
    else if (!unknown.includes(name)) unknown.push(name);
  });
  return { keys: INPUTS.map((i) => i.key).filter((k) => keys.has(k)), unknown };
}

/** 고른 서식에 필요한데 비어 있는 입력(작성일·진료 종료일은 비어도 됨) */
function missingInputs(keys, input = {}) {
  return keys.filter((k) => !OPTIONAL_KEYS.includes(k) && !String(input[k] == null ? '' : input[k]).trim());
}

module.exports = { INPUTS, INPUT_BY_KEY, INPUT_GROUPS, SAMPLE_INPUT, FIELDS, FIELD_BY_ID, fieldOf, unknownFields, sourceLabel, fieldValues, inputsFor, missingInputs, parseDate, shortDate, carePeriod, rrnParts, checkRrn, rrnSummary };
