// field-dictionary.js
// 서식의 머리글(라벨) ↔ 입력 항목 사전과, 입력값에서 서식에 넣을 값을 만드는 규칙.
// 같은 "성명"이라도 "등록 대상자" 묶음이면 환자, "신청인(보호자 등)" 묶음이면 신청인이다 — 묶음 머리글로 구분한다.
// AI 없이 정해진 규칙으로만 판단한다. 화면에 기대지 않아 node로 단위 테스트한다.

/** 입력 항목. section: patient(등록 대상자) · applicant(신청인) · diagnosis(진단 정보) · org(기관 정보, 설정에 저장) · auto(자동) */
const FIELDS = [
  { key: 'patient.name', section: 'patient', label: '성명', required: true },
  { key: 'patient.rrn', section: 'patient', label: '주민등록번호', required: true },
  { key: 'patient.birth', section: 'auto', label: '생년월일', derived: true },
  { key: 'patient.sex', section: 'auto', label: '성별', derived: true },
  { key: 'patient.age', section: 'auto', label: '나이', derived: true },
  { key: 'patient.addr', section: 'patient', label: '주소' },
  { key: 'patient.phone', section: 'patient', label: '전화번호' },
  { key: 'patient.mobile', section: 'patient', label: '휴대폰번호' },
  { key: 'patient.email', section: 'patient', label: '이메일주소' },
  { key: 'patient.household', section: 'patient', label: '가구원수', unit: '명' },
  { key: 'patient.bank', section: 'patient', label: '금융기관' },
  { key: 'patient.holder', section: 'patient', label: '예금주', hint: '비우면 환자 성명' },
  { key: 'patient.account', section: 'patient', label: '계좌번호' },
  { key: 'applicant.name', section: 'applicant', label: '성명' },
  { key: 'applicant.rrn', section: 'applicant', label: '주민등록번호' },
  { key: 'applicant.relation', section: 'applicant', label: '관계(환자와의)' },
  { key: 'applicant.phone', section: 'applicant', label: '전화번호' },
  { key: 'applicant.addr', section: 'applicant', label: '주소' },
  { key: 'diagnosis.name', section: 'diagnosis', label: '상병명' },
  { key: 'diagnosis.code', section: 'diagnosis', label: '상병코드' },
  { key: 'diagnosis.firstDate', section: 'diagnosis', label: '최초진단일' },
  { key: 'org.name', section: 'org', label: '진단기관' },
  { key: 'org.phone', section: 'org', label: '기관 전화번호' },
  { key: 'org.addr', section: 'org', label: '기관 소재지' },
  { key: 'auto.today', section: 'auto', label: '작성일', derived: true },
  { key: 'auto.signer', section: 'auto', label: '서명란 이름(신청인·동의자·민원인)', derived: true },
  { key: 'auto.signerRelation', section: 'auto', label: '신청인과의 관계(대리 신청일 때)', derived: true },
];
const FIELD_BY_KEY = new Map(FIELDS.map((f) => [f.key, f]));

const SECTION_LABELS = { patient: '등록 대상자(환자)', applicant: '신청인(보호자 등)', diagnosis: '진단 정보', org: '기관 정보', auto: '자동' };

/** 비교용: 공백·가운뎃점 없애고, 괄호 안 설명은 뺀다("가구원수 (환자 포함)" → "가구원수") */
function normalizeLabel(s) {
  return String(s)
    .replace(/[(（][^)）]*[)）]/g, '')
    .replace(/[\s·‧ㆍ:：]/g, '')
    .toLowerCase();
}

/** 묶음 머리글("등록 대상자", "신청인 (보호자 등)", "진단정보") → patient | applicant | diagnosis | null */
function groupKind(text) {
  const n = String(text || '').replace(/\s+/g, '');
  if (!n) return null;
  if (/(신청인|보호자|대리인|대리신청|제출인)/.test(n)) return 'applicant';
  if (/(대상자|환자|본인|수급자|수혜자)/.test(n)) return 'patient';
  if (/(진단|상병|진료정보|의료기관)/.test(n)) return 'diagnosis';
  return null;
}

// 라벨 → 항목 종류(사람·기관 구분 전). [정규식, 종류, 확인 필요 이유]
const LABEL_RULES = [
  [/^(성명|이름|환자명|대상자성명|성명한글)$/, 'name'],
  [/^(주민등록번호|주민번호)$/, 'rrn'],
  [/^생년월일$/, 'birth'],
  [/^성별$/, 'sex'],
  [/^(나이|연령|만나이)$/, 'age'],
  [/^(주소|거주지|주소지|현주소)$/, 'addr'],
  [/^(전화번호|전화|자택전화|집전화|자택전화번호)$/, 'phone'],
  [/^(휴대폰번호|휴대폰|휴대전화|휴대전화번호|핸드폰|핸드폰번호|이동전화)$/, 'mobile'],
  [/^연락처$/, 'mobile', '“연락처”는 휴대폰·전화 둘 다 될 수 있습니다'],
  [/^(이메일주소|이메일|전자우편|e-?mail)$/, 'email'],
  [/^가구원수/, 'household'],
  [/^(금융기관|은행|은행명|금융기관명)$/, 'bank'],
  [/^예금주$/, 'holder'],
  [/^계좌번호$/, 'account'],
  [/^(관계|환자와의관계|대상자와의관계|본인과의관계)$/, 'relation'],
  [/^(상병명|진단명|병명|질병명)$/, 'diagnosis.name'],
  [/^(상병코드|질병코드|상병분류기호|질병분류기호)$/, 'diagnosis.code'],
  [/^(최초진단일|진단일|진단일자|최초진단일자)$/, 'diagnosis.firstDate'],
  [/^(진단기관|의료기관|의료기관명|병원명|진단기관명|요양기관명)$/, 'org.name'],
  [/^(소재지|의료기관소재지)$/, 'addrOrOrg'],
  [/^(신청인과의관계)$/, 'auto.signerRelation'],
  [/^(신청인|대리신청인|동의자|민원인|신청자|제출인)$/, 'auto.signer'],
];

/**
 * 라벨(머리글) 하나를 입력 항목으로. group: groupKind() 결과(없으면 null).
 * 반환: { field, confidence: 'sure'|'check', why } 또는 null(맞는 항목 없음)
 */
function matchLabel(label, group) {
  const n = normalizeLabel(label);
  if (!n) return null;
  const rule = LABEL_RULES.find(([re]) => re.test(n));
  if (!rule) return null;
  const [, kind, why] = rule;
  const out = (field, w) => ({ field, confidence: w ? 'check' : 'sure', why: w || '' });
  if (kind.includes('.')) return out(kind, why);
  if (group === 'diagnosis') {
    if (kind === 'phone' || kind === 'mobile') return out('org.phone', why);
    if (kind === 'addr' || kind === 'addrOrOrg') return out('org.addr', why);
  }
  if (kind === 'addrOrOrg') return out(group === 'applicant' ? 'applicant.addr' : 'org.addr');
  if (group === 'applicant') {
    if (kind === 'mobile') return out('applicant.phone', why);
    if (['name', 'rrn', 'relation', 'phone', 'addr'].includes(kind)) return out(`applicant.${kind}`, why);
  }
  if (kind === 'relation') return out('applicant.relation', '묶음 없이 “관계”만 있어 신청인의 관계로 보았습니다');
  return out(`patient.${kind}`, why);
}

/* ---------- 입력값 → 서식에 넣을 값 ---------- */

/** 주민등록번호 → { birth: Date 부분, sex: '남'|'여', age } (형식이 틀리면 null) */
function parseRrn(rrn) {
  const m = /^(\d{2})(\d{2})(\d{2})-?(\d)(\d{6})$/.exec(String(rrn || '').trim());
  if (!m) return null;
  const g = Number(m[4]);
  const century = [1, 2, 5, 6].includes(g) ? 1900 : [3, 4, 7, 8].includes(g) ? 2000 : [9, 0].includes(g) ? 1800 : null;
  if (!century) return null;
  const y = century + Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return { y, m: mo, d, sex: g % 2 === 1 ? '남' : '여' };
}

function ageOn(birth, today) {
  let age = today.getFullYear() - birth.y;
  if (today.getMonth() + 1 < birth.m || (today.getMonth() + 1 === birth.m && today.getDate() < birth.d)) age -= 1;
  return age;
}

/**
 * 화면 입력(input)에서 서식에 넣을 값(항목 키 → 글자)을 만든다.
 * input: { values: { 'patient.name': ..., ... }, org: { name, phone, addr }, applicantMode: 'self'|'proxy', today: Date }
 * - 생년월일·성별·나이: 주민등록번호에서
 * - 예금주: 비우면 환자 성명
 * - 서명란 이름: 보호자·대리인 신청이면 신청인 성명, 아니면 환자 성명
 * - 신청인과의 관계: 대리 신청일 때만 신청인의 관계
 * - 주민등록번호는 "000000-0000000" 모양으로 맞춘다
 */
function resolveValues({ values = {}, org = {}, applicantMode = 'self', today = new Date() } = {}) {
  const v = {};
  Object.entries(values).forEach(([k, x]) => {
    v[k] = x === undefined || x === null ? '' : String(x).trim();
  });
  const fmtRrn = (s) => {
    const m = /^(\d{6})-?(\d{7})$/.exec(s || '');
    return m ? `${m[1]}-${m[2]}` : s || '';
  };
  v['patient.rrn'] = fmtRrn(v['patient.rrn']);
  v['applicant.rrn'] = fmtRrn(v['applicant.rrn']);
  const birth = parseRrn(v['patient.rrn']);
  v['patient.birth'] = birth ? `${birth.y}. ${birth.m}. ${birth.d}.` : '';
  v['patient.sex'] = birth ? birth.sex : '';
  v['patient.age'] = birth ? String(ageOn(birth, today)) : '';
  if (!v['patient.holder']) v['patient.holder'] = v['patient.name'] || '';
  v['org.name'] = org.name || v['org.name'] || '';
  v['org.phone'] = org.phone || v['org.phone'] || '';
  v['org.addr'] = org.addr || v['org.addr'] || '';
  const proxy = applicantMode === 'proxy';
  if (!proxy) ['applicant.name', 'applicant.rrn', 'applicant.relation', 'applicant.phone', 'applicant.addr'].forEach((k) => (v[k] = ''));
  v['auto.signer'] = proxy && v['applicant.name'] ? v['applicant.name'] : v['patient.name'] || '';
  v['auto.signerRelation'] = proxy ? v['applicant.relation'] || '' : '';
  v['auto.today'] = `${today.getFullYear()}. ${today.getMonth() + 1}. ${today.getDate()}.`;
  v.__today = { y: today.getFullYear(), m: today.getMonth() + 1, d: today.getDate() };
  return v;
}

module.exports = { FIELDS, FIELD_BY_KEY, SECTION_LABELS, normalizeLabel, groupKind, matchLabel, parseRrn, resolveValues };
