// form-library.js
// 서식 보관함: 등록한 한글 서식(원본 .hwp/.hwpx와 한글로 바꾼 .hwpx), 서식마다 제출 기관·지원 종류·신청 구분
// (세대주/세대원), 들어 있는 메일머지 필드와 서류 목차, 필드 연결(매핑)을 이 PC에 저장한다.
// 환자 인적사항은 여기 저장하지 않는다(서식 파일과 설정만). 화면(DOM)에 기대지 않아 node로 단위 테스트한다.
//
// 저장 위치: <dir>/index.json(목록) · <dir>/files/<id>.hwp|.hwpx(원본 사본) · <dir>/files/<id>.conv.hwpx(변환본)

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const PizZip = require('pizzip');
const { fieldOf, unknownFields } = require('./merge-fields.js');
const { readHwpFields, documentOutline } = require('../engine/hwp-reader.js');
const { modelFromHwpx } = require('../engine/form-model.js');
const { recognizePlaceholders } = require('../engine/form-recognize.js');

const KINDS = ['긴급복지(국가형)', '경기도형 긴급복지', '서울형 긴급복지', '광주형 긴급지원', '암환자 의료비', '기타'];
const APPLICANTS = ['세대주', '세대원', '공통'];
const COMMON_AGENCY = '공통 서식';
const AGENCY_SUFFIX = /(시청|구청|군청|도청|보건소|보건지소|행정복지센터|주민센터)$/;
// 지원 종류: 이름(폴더 포함)에 든 말로, 위에서부터 먼저 맞는 것
const KIND_RULES = [
  [/암환자/, '암환자 의료비'],
  [/서울형/, '서울형 긴급복지'],
  [/광주형/, '광주형 긴급지원'],
  [/경기도형/, '경기도형 긴급복지'],
  [/긴급/, '긴급복지(국가형)'],
];
const SEOUL_AGENCY = '서울특별시';

/* ---------- 파일 이름으로 분류 ---------- */

/** "(세대주)"/"(세대원)" 표시 → 구분(없으면 null) */
const applicantIn = (text) => (/\(세대주\)/.test(text) ? '세대주' : /\(세대원\)/.test(text) ? '세대원' : null);

/**
 * 파일 경로(폴더 포함)로 서식을 분류한다. 예: "1. (세대주)…/★포천시청-긴급지원(의료비) 신청서(세대주).hwp"
 * → { agency: '포천시청', kind: '긴급복지(국가형)', applicant: '세대주', favorite: true, title: '긴급지원(의료비) 신청서' }
 * - 기관: 이름 맨 앞 "…시청·구청·군청·보건소" 다음에 "-"나 "_"가 오면 그 앞까지. 없으면 서울형은 "서울특별시", 그 밖은 "공통 서식"
 * - 종류: KIND_RULES · 구분: 파일 이름, 없으면 폴더 이름의 (세대주)/(세대원), 둘 다 없으면 공통 · ★로 시작하면 자주 쓰는 서식
 */
function classifyFileName(filePath) {
  const parts = String(filePath).split(/[\\/]/);
  const base = parts.pop().replace(/\.(hwpx?|hwpt)$/i, '');
  const whole = `${parts.join('/')}/${base}`;
  const favorite = /^\s*★/.test(base);
  let rest = base.replace(/^\s*★\s*/, '');

  let agency = '';
  const m = /^(.{2,20}?)\s*[-_]\s*(.+)$/.exec(rest);
  if (m && AGENCY_SUFFIX.test(m[1].trim())) {
    agency = m[1].trim();
    rest = m[2];
  }
  const kind = (KIND_RULES.find(([re]) => re.test(whole)) || [null, '기타'])[1];
  if (!agency) agency = kind === '서울형 긴급복지' ? SEOUL_AGENCY : COMMON_AGENCY;
  const applicant = applicantIn(base) || applicantIn(whole) || '공통';
  const title =
    rest
      .replace(/\s*\((세대주|세대원)\)\s*/g, ' ')
      .replace(/\s*[-_]?\s*\(?메일머지\)?\s*$/, '')
      .replace(/^\s*공통\s*[_-]\s*/, '')
      .replace(/\s+/g, ' ')
      .trim() || base;
  return { agency, kind, applicant, favorite, title };
}

/* ---------- 주소로 제출 기관 추천 ---------- */

const REGION_SUFFIX = '(?:특별시|광역시|특별자치시|특별자치도|시|군|구)';

/** 주소의 시·군·구 이름들("경기도 양주시 백석읍" → ["양주시"], "서울특별시 노원구" → ["서울특별시", "노원구"]) */
function regionsOf(address) {
  return [...String(address || '').matchAll(new RegExp(`([가-힣]{1,10}${REGION_SUFFIX})(?=\\s|,|$)`, 'g'))].map((m) => m[1]).filter((r) => !/^(경기|강원|충청|전라|경상|제주)?도$/.test(r));
}

/**
 * 주소에 맞는 기관 이름들(작은 지역부터). "양주시" → "양주시청", "노원구" → "노원구청", "광산구" → "광주 광산구청",
 * "동두천시" → "동두천시청"·"동두천 보건소", 서울 주소 → "서울특별시"(서울형). 맞는 것이 없으면 빈 목록.
 */
function suggestAgencies(address, agencyNames) {
  const out = [];
  const push = (name) => {
    if (!out.includes(name)) out.push(name);
  };
  regionsOf(address)
    .reverse()
    .forEach((region) => {
      const stem = region.replace(new RegExp(`${REGION_SUFFIX}$`), '');
      if (!stem) return;
      // "양주시" ⊂ "양주시청", 또는 기관 이름이 "동두천"+(시·군·구·보건소)로 시작
      const startsWithStem = new RegExp(`^${stem}(시|군|구|보건)`);
      agencyNames.forEach((name) => {
        const compact = name.replace(/\s/g, '');
        if (compact.includes(region) || (stem.length >= 2 && startsWithStem.test(compact))) push(name);
      });
      if (/^서울/.test(region)) agencyNames.filter((n) => /^서울/.test(n)).forEach(push);
    });
  return out;
}

/* ---------- 서식 파일 읽기 ---------- */

/**
 * 서식 파일의 메일머지 필드와 서류 목차. 반환: { format: 'hwp'|'hwpx', fields: [{ name, count }], documents: [{ mark, title }] }
 * .hwp는 한글 없이 직접, .hwpx는 서식 엔진 모델로 읽는다({{필드}} 글자).
 */
function inspectFormFile(filePath) {
  const buffer = fs.readFileSync(filePath);
  if (path.extname(filePath).toLowerCase() === '.hwp') return { format: 'hwp', ...readHwpFields(buffer) };
  const model = modelFromHwpx(buffer);
  const counts = new Map();
  recognizePlaceholders(model)
    .filter((s) => !/^[#/^.]/.test(s.label)) // {{#반복}} 같은 제어 태그는 필드가 아니다
    .forEach((s) => counts.set(s.label, (counts.get(s.label) || 0) + 1));
  return { format: 'hwpx', fields: [...counts].map(([name, count]) => ({ name, count })), documents: documentOutline(model.paragraphs.map((p) => p.text)) };
}

/** 폴더 안의 서식 파일(.hwp·.hwpx·.hwpt) 경로들(하위 폴더까지). 한글 임시 파일(~$…)·숨은 파일은 뺀다 */
function findFormFiles(target) {
  if (!fs.statSync(target).isDirectory()) return /\.(hwpx?|hwpt)$/i.test(target) ? [target] : [];
  return fs
    .readdirSync(target, { withFileTypes: true })
    .filter((e) => !e.name.startsWith('.') && !e.name.startsWith('~$'))
    .sort((a, b) => a.name.localeCompare(b.name, 'ko'))
    .flatMap((e) => findFormFiles(path.join(target, e.name)));
}

/**
 * 등록 확인 창의 줄들: 파일·폴더 경로들 → [{ file, agency, kind, applicant, favorite, title, fields, documents, unknown, error, dup, include }]
 * 읽지 못한 파일과 이미 등록한 서식(같은 원본 이름·기관·구분)은 처음에 고르지 않는다.
 */
function prepareImport(paths, existing = []) {
  const files = [...new Set(paths.flatMap((p) => findFormFiles(p)))];
  return files.map((file) => {
    const info = classifyFileName(file);
    let inspected = { fields: [], documents: [] };
    let error = '';
    try {
      inspected = inspectFormFile(file);
    } catch (err) {
      error = err.message;
    }
    const dup = existing.some((e) => e.sourceName === path.basename(file) && e.agency === info.agency && e.applicant === info.applicant);
    return { file, ...info, fields: inspected.fields, documents: inspected.documents, unknown: unknownFields(inspected.fields), error, dup, include: !error && !dup };
  });
}

/* ---------- 서식 상태 · 순서 ---------- */

/** 필드 연결 기본값: 사전에 있는 필드는 자동, 없는 필드는 출력할 때 묻기 */
function defaultMapping(fields) {
  return Object.fromEntries(fields.map((f) => [f.name, fieldOf(f.name) ? { type: 'auto' } : { type: 'ask' }]));
}

/**
 * 서식 상태: { ready(출력 목록에서 고를 수 있는지), needsConvert, label, problems }
 * - 사전에 없는 필드를 아직 검토하지 않았으면 확인 필요
 * - 한글로 바꾼 .hwpx가 없으면(.hwp를 아직 변환 못 함) 변환 전 — 출력할 때 한글이 있으면 그때 바꾼다
 */
function entryStatus(entry) {
  const unknown = unknownFields(entry.fields);
  const unchecked = unknown.length > 0 && !entry.reviewed;
  const problems = [];
  if (unchecked) problems.push(`사전에 없는 필드 ${unknown.length}개 확인 필요(${unknown.join(', ')})`);
  if (entry.fields.length === 0) problems.push('메일머지 필드가 없어 자동 인식으로 채웁니다(시험 출력으로 확인 권장)');
  const needsConvert = !entry.hwpx;
  return { ready: !unchecked, needsConvert, label: unchecked ? '확인 필요' : needsConvert ? '한글 변환 전' : '사용 중', problems };
}

const commonLast = (a, b) => (a === COMMON_AGENCY) - (b === COMMON_AGENCY);

/** 서식 관리 목록 순서: 기관(공통 서식은 맨 뒤) → 지원 종류 → 구분 → ★ 먼저 */
function sortForManager(entries) {
  return entries
    .slice()
    .sort((a, b) => commonLast(a.agency, b.agency) || a.agency.localeCompare(b.agency, 'ko') || KINDS.indexOf(a.kind) - KINDS.indexOf(b.kind) || a.applicant.localeCompare(b.applicant, 'ko') || b.favorite - a.favorite);
}

/* ---------- 보관함 ---------- */

const newId = () => crypto.randomBytes(6).toString('hex');
const ID = /^[0-9a-f]{12}$/;
/** 서식 하나의 보관 파일 이름들(원본·변환본, 겹치면 한 번) */
const entryFiles = (e) => [...new Set([e.file, e.hwpx].filter(Boolean))];

function writeJsonAtomic(file, data) {
  const tmp = `${file}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
  fs.renameSync(tmp, file);
}

/**
 * 백업에서 읽은 항목을 알려진 칸만, 알맞은 모양으로. 파일 이름은 그 항목의 id로 시작해야 한다(다른 서식 파일을
 * 덮거나 보관 폴더 밖을 가리키지 않게). 맞지 않으면 null
 */
function normalizeEntry(raw) {
  if (!raw || !ID.test(raw.id)) return null;
  const ownFile = (n) => typeof n === 'string' && new RegExp(`^${raw.id}(\\.conv)?\\.hwpx?$`).test(n);
  if (!ownFile(raw.file) || (raw.hwpx != null && !ownFile(raw.hwpx))) return null;
  const text = (v, fallback) => (typeof v === 'string' && v.trim() ? v.trim() : fallback);
  const list = (v) => (Array.isArray(v) ? v : []);
  return {
    id: raw.id,
    title: text(raw.title, raw.id),
    agency: text(raw.agency, COMMON_AGENCY),
    kind: KINDS.includes(raw.kind) ? raw.kind : '기타',
    applicant: APPLICANTS.includes(raw.applicant) ? raw.applicant : '공통',
    favorite: !!raw.favorite,
    sourceName: text(raw.sourceName, raw.file),
    file: raw.file,
    hwpx: raw.hwpx || null,
    fields: list(raw.fields).filter((f) => f && typeof f.name === 'string').map((f) => ({ name: f.name, count: Number(f.count) || 1 })),
    documents: list(raw.documents).filter((d) => d && typeof d === 'object').map((d) => ({ mark: String(d.mark || ''), title: String(d.title || '') })),
    mapping: raw.mapping && typeof raw.mapping === 'object' ? raw.mapping : {},
    reviewed: !!raw.reviewed,
    addedAt: text(raw.addedAt, new Date().toISOString()),
    updatedAt: text(raw.updatedAt, new Date().toISOString()),
  };
}

/** dir 아래의 서식 보관함. 반환하는 객체의 메서드는 모두 바로 디스크에 반영한다 */
function openLibrary(dir) {
  const filesDir = path.join(dir, 'files');
  const indexFile = path.join(dir, 'index.json');
  fs.mkdirSync(filesDir, { recursive: true });
  let entries = [];
  let loadError = null;
  if (fs.existsSync(indexFile)) {
    try {
      const parsed = JSON.parse(fs.readFileSync(indexFile, 'utf8'));
      entries = Array.isArray(parsed.entries) ? parsed.entries : [];
    } catch (err) {
      loadError = err;
      // 망가진 목록은 지우지 않고 옆에 남겨 둔다(서식 파일은 files 폴더에 그대로 있음)
      fs.copyFileSync(indexFile, `${indexFile}.broken-${Date.now()}`);
    }
  }
  const save = () => writeJsonAtomic(indexFile, { version: 1, entries });
  const find = (id) => {
    const e = entries.find((x) => x.id === id);
    if (!e) throw new Error('서식을 찾을 수 없습니다. 목록을 새로 고친 뒤 다시 시도하세요.');
    return e;
  };
  const touch = (e) => {
    e.updatedAt = new Date().toISOString();
    save();
    return e;
  };

  return {
    dir,
    loadError,
    list: () => entries.slice(),
    get: find,
    /** 원본 파일의 보관 사본 경로 */
    sourcePath: (entry) => path.join(filesDir, entry.file),
    /** 한글로 바꾼(또는 원래 .hwpx인) 채울 파일 경로. 없으면 null */
    hwpxPath: (entry) => (entry.hwpx ? path.join(filesDir, entry.hwpx) : null),
    /**
     * 서식 등록. source: 원본 파일 경로. info: { agency, kind, applicant, favorite, title, fields, documents, mapping }
     * .hwpx는 그대로 채울 파일로 쓰고, .hwp는 원본만 두었다가 한글로 바꾸면 setConverted로 넣는다.
     */
    add(source, info) {
      const ext = path.extname(source).toLowerCase() === '.hwp' ? '.hwp' : '.hwpx';
      const id = newId();
      const file = `${id}${ext}`;
      fs.copyFileSync(source, path.join(filesDir, file));
      const now = new Date().toISOString();
      const fields = info.fields || [];
      const entry = {
        id,
        title: info.title,
        agency: info.agency,
        kind: info.kind,
        applicant: info.applicant,
        favorite: !!info.favorite,
        sourceName: path.basename(source),
        file,
        hwpx: ext === '.hwpx' ? file : null,
        fields,
        documents: info.documents || [],
        mapping: info.mapping || defaultMapping(fields),
        reviewed: false,
        addedAt: now,
        updatedAt: now,
      };
      entries.push(entry);
      save();
      return entry;
    },
    /** 한글로 바꾼 .hwpx 파일을 넣는다(converted: 바꾼 파일 경로) */
    setConverted(id, converted) {
      const e = find(id);
      e.hwpx = `${e.id}.conv.hwpx`;
      fs.copyFileSync(converted, path.join(filesDir, e.hwpx));
      return touch(e);
    },
    /** 이름·기관·종류·구분·★·매핑·검토 여부·필드 목록 고치기 */
    update(id, patch) {
      const e = find(id);
      ['title', 'agency', 'kind', 'applicant', 'favorite', 'mapping', 'reviewed', 'fields', 'documents'].forEach((k) => {
        if (patch[k] !== undefined) e[k] = patch[k];
      });
      return touch(e);
    },
    remove(id) {
      const e = find(id);
      entries = entries.filter((x) => x !== e);
      save();
      entryFiles(e).forEach((n) => fs.rmSync(path.join(filesDir, n), { force: true }));
    },
    /** 기관 목록: [{ name, count }] — 공통 서식은 맨 뒤 */
    agencies() {
      const counts = new Map();
      entries.forEach((e) => counts.set(e.agency, (counts.get(e.agency) || 0) + 1));
      return [...counts].map(([name, count]) => ({ name, count })).sort((a, b) => commonLast(a.name, b.name) || a.name.localeCompare(b.name, 'ko'));
    },
    /** 백업 파일(.zip 버퍼): 목록과 서식 파일들 */
    exportBackup() {
      const zip = new PizZip();
      zip.file('index.json', JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), entries }, null, 2));
      entries.forEach((e) => entryFiles(e).forEach((n) => zip.file(`files/${n}`, fs.readFileSync(path.join(filesDir, n)))));
      return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    },
    /** 백업 가져오기: 같은 id는 백업 것으로 바꾸고, 없던 것은 더한다. 모양이 맞지 않는 항목은 건너뛴다. 반환: { added, replaced } */
    importBackup(buffer) {
      let zip;
      let data;
      try {
        zip = new PizZip(buffer);
        data = JSON.parse(zip.file('index.json').asText());
      } catch (err) {
        throw new Error('서식 백업 파일이 아닙니다(또는 손상되었습니다).');
      }
      let added = 0;
      let replaced = 0;
      (Array.isArray(data.entries) ? data.entries : []).forEach((raw) => {
        const e = normalizeEntry(raw);
        if (!e || !entryFiles(e).every((n) => zip.file(`files/${n}`))) return;
        entryFiles(e).forEach((n) => fs.writeFileSync(path.join(filesDir, n), zip.file(`files/${n}`).asNodeBuffer()));
        const i = entries.findIndex((x) => x.id === e.id);
        if (i >= 0) {
          entries[i] = e;
          replaced += 1;
        } else {
          entries.push(e);
          added += 1;
        }
      });
      save();
      return { added, replaced };
    },
  };
}

module.exports = {
  KINDS,
  APPLICANTS,
  COMMON_AGENCY,
  classifyFileName,
  regionsOf,
  suggestAgencies,
  inspectFormFile,
  findFormFiles,
  prepareImport,
  defaultMapping,
  entryStatus,
  sortForManager,
  openLibrary,
};
