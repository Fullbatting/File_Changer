// print-job.js
// 서류 출력: 등록한 서식(.hwpx)에 환자 인적사항을 채워 한글로 바로 인쇄하고, 채운 임시 파일은 곧바로 지운다.
// - 메일머지 필드가 있는 서식: {{이름}}·{{1}}~{{13}}·{{월1}} … 자리에 merge-fields.js의 값을 넣는다.
//   "2026년 {{월1}}{{월2}}월"처럼 연도가 미리 찍혀 있으면 그 연도도 작성일의 연도로 맞춘다.
// - 필드가 없는 서식: 서식 엔진의 자동 인식(form-recognize.js)으로 채운다.
// - .hwp 서식은 한글(hancom.js)로 .hwpx로 바꿔 두고 쓴다(등록 때 못 바꿨으면 처음 출력할 때).
// 인적사항은 메모리에만 두고, 로그에도 값을 남기지 않는다.

const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { modelFromHwpx } = require('../engine/form-model.js');
const { recognize, recognizePlaceholders } = require('../engine/form-recognize.js');
const { writePlan, planFill } = require('../engine/form-fill.js');
const { resolveValues } = require('../engine/field-dictionary.js');
const { fieldValues, shortDate, fieldOf } = require('./merge-fields.js');
const { runHancom } = require('./hancom.js');

const TEMP_PREFIX = 'form-autofill-';

/** 필드 값: 서식의 매핑(auto 사전 값 · fixed 고정 글자 · ask 출력 때 입력 · blank 비워 둠) */
function valueFor(name, mapping, values, asks) {
  const m = mapping[name] || (fieldOf(name) ? { type: 'auto' } : { type: 'ask' });
  if (m.type === 'fixed') return m.text || '';
  if (m.type === 'blank') return '';
  if (m.type === 'ask') return String(asks[name] == null ? '' : asks[name]);
  return values[name] || '';
}

/**
 * 채울 .hwpx 버퍼를 만든다. entry: 보관함 항목(mapping) · input: 화면 입력 · asks: { 필드 이름: 글자 }(출력 때 입력한 것)
 * 반환: { buffer, mode: 'fields'|'auto', filled(채운 자리 수) }
 */
function fillForm(entry, hwpxBuffer, input, { asks = {}, today = new Date() } = {}) {
  const model = modelFromHwpx(hwpxBuffer);
  const spots = recognizePlaceholders(model);
  if (spots.length === 0) return fillByRecognition(hwpxBuffer, model, input, today);
  const plan = fieldPlan(entry, model, spots, input, { asks, today });
  return { buffer: writePlan(hwpxBuffer, model, plan), mode: 'fields', filled: spots.length };
}

/**
 * 메일머지 필드 자리마다 넣을 글자: [{ spotId, paragraphId, start, end, text }].
 * "2026년 {{월1}}{{월2}}월"처럼 연도가 미리 찍혀 있으면 그 네 자리도 작성일의 연도로(yearOf: 그 월1 자리)
 */
function fieldPlan(entry, model, spots, input, { asks = {}, today = new Date() } = {}) {
  const names = [...new Set(spots.map((s) => s.label))];
  const values = fieldValues(names, input, { today });
  const plan = spots.map((s) => ({ spotId: s.id, paragraphId: s.paragraphId, start: s.range.start, end: s.range.end, text: valueFor(s.label, entry.mapping || {}, values, asks) }));
  const year = String(values.__written.y);
  spots
    .filter((s) => fieldOf(s.label) && fieldOf(s.label).id === 'month1')
    .forEach((s) => {
      const text = model.paragraphById.get(s.paragraphId).text;
      const m = /(\d{4})(\s*년\s*)$/.exec(text.slice(0, s.range.start));
      if (m && m[1] !== year) plan.push({ spotId: `${s.id}:year`, paragraphId: s.paragraphId, start: m.index, end: m.index + 4, text: year, yearOf: s });
    });
  return plan;
}

/**
 * 한글의 "모두 바꾸기"에 넘길 [찾을 글자, 바꿀 글자] 목록(메일머지 필드 서식만, 아니면 null).
 * 찾을 글자는 서식에 적힌 그대로("{{이름}}"). 연도를 맞출 자리는 "2026년 {{월1}}"을 통째로 찾아 먼저 바꾼다.
 */
function replacementsFor(entry, hwpxBuffer, input, opts = {}) {
  const model = modelFromHwpx(hwpxBuffer);
  const spots = recognizePlaceholders(model);
  if (spots.length === 0) return null;
  const plan = fieldPlan(entry, model, spots, input, opts);
  const textOf = (pid) => model.paragraphById.get(pid).text;
  const pairs = new Map();
  plan
    .filter((e) => e.yearOf)
    .forEach((e) => {
      const t = textOf(e.paragraphId);
      const month = plan.find((x) => x.spotId === e.yearOf.id);
      pairs.set(t.slice(e.start, e.yearOf.range.end), e.text + t.slice(e.end, e.yearOf.range.start) + month.text);
    });
  plan
    .filter((e) => !e.yearOf)
    .forEach((e) => {
      const find = textOf(e.paragraphId).slice(e.start, e.end);
      if (!pairs.has(find)) pairs.set(find, e.text);
    });
  return [...pairs.entries()];
}

/** 필드 없는 서식: 자동 인식(form-recognize.js) + 그 엔진의 입력 항목 사전(field-dictionary.js)으로 */
function fillByRecognition(hwpxBuffer, model, input, today) {
  const spots = recognize(model);
  const values = resolveValues({
    values: {
      'patient.name': input.name,
      'patient.rrn': input.rrn,
      'patient.addr': input.addr,
      'patient.phone': input.phone,
      'patient.mobile': input.phone,
      'diagnosis.name': input.disease,
      'diagnosis.code': input.diseaseCode,
      'diagnosis.firstDate': input.diagnosisDate ? shortDate(input.diagnosisDate) : '',
    },
    applicantMode: 'self',
    today,
  });
  const plan = planFill(model, spots, { values, choices: {} });
  return { buffer: writePlan(hwpxBuffer, model, plan), mode: 'auto', filled: plan.length };
}

/** 임시 폴더(남은 임시 파일은 시작할 때 cleanupTemp로 지운다) */
const tempDir = () => {
  const dir = path.join(os.tmpdir(), 'form-autofill');
  fs.mkdirSync(dir, { recursive: true });
  return dir;
};
const tempFile = (ext) => path.join(tempDir(), `${TEMP_PREFIX}${crypto.randomBytes(8).toString('hex')}${ext}`);

/** 지난번에 지우지 못한 임시 파일(프로그램이 갑자기 꺼진 경우 등)을 지운다. 반환: 지운 개수 */
function cleanupTemp() {
  let n = 0;
  const dir = tempDir();
  fs.readdirSync(dir)
    .filter((f) => f.startsWith(TEMP_PREFIX))
    .forEach((f) => {
      try {
        fs.rmSync(path.join(dir, f), { force: true });
        n += 1;
      } catch (e) {
        /* 한글이 아직 쥐고 있으면 다음에 지운다 */
      }
    });
  return n;
}

/**
 * .hwp 서식을 한글로 .hwpx로 바꿔 보관함에 넣는다(이미 있으면 그대로). 반환: 채울 .hwpx 경로
 * hancom: runHancom 대신 쓸 함수(테스트용)
 */
async function ensureHwpx(library, entry, { hancom = runHancom } = {}) {
  const existing = library.hwpxPath(entry);
  if (existing && fs.existsSync(existing)) return existing;
  const out = tempFile('.hwpx');
  try {
    const r = await hancom('convert', { input: library.sourcePath(entry), output: out });
    if (!r.ok) throw new Error(`한글로 .hwpx 변환을 하지 못했습니다: ${r.error || '알 수 없는 오류'}`);
    return library.hwpxPath(library.setConverted(entry.id, out));
  } finally {
    fs.rmSync(out, { force: true });
  }
}

/**
 * 고른 서식들을 채워 인쇄한다. 서식마다 결과를 모아 돌려준다(한 서식이 실패해도 나머지는 계속).
 * options: { copies, printer, asks, today, hancom, onProgress(i, total, entry) }
 * 반환: [{ id, title, ok, error }]
 */
async function printForms(library, entries, input, { copies = 1, printer = '', asks = {}, today = new Date(), hancom = runHancom, onProgress } = {}) {
  const results = [];
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    if (onProgress) onProgress(i, entries.length, entry);
    const file = tempFile('.hwpx');
    try {
      const source = await ensureHwpx(library, entry, { hancom });
      const buffer = fs.readFileSync(source);
      fs.writeFileSync(file, fillForm(entry, buffer, input, { asks, today }).buffer);
      let r = await hancom('print', { file, copies, printer });
      let via = 'file';
      if (!r.ok && /열지 못했습니다/.test(r.error || '')) {
        // 채운 파일을 한글이 못 열면: 원본 서식을 한글로 열어 {{필드}}를 한글이 직접 바꾼 뒤 인쇄(파일은 저장하지 않음)
        const replacements = replacementsFor(entry, buffer, input, { asks, today });
        if (replacements) {
          const r2 = await hancom('printReplace', { file: library.sourcePath(entry), replacements, copies, printer });
          if (!r2.ok) throw new Error(`채운 서식 파일을 한글이 열지 못했고, 원본을 한글로 직접 채워 인쇄하는 것도 실패했습니다: ${r2.error || '알 수 없는 오류'}`);
          r = r2;
          via = 'hancom';
        }
      }
      if (!r.ok) throw new Error(r.error || '한글이 인쇄를 시작하지 못했습니다');
      results.push({ id: entry.id, title: entry.title, ok: true, error: '', via });
    } catch (err) {
      results.push({ id: entry.id, title: entry.title, ok: false, error: err.message || String(err) });
    } finally {
      fs.rmSync(file, { force: true }); // 인적사항이 든 파일은 남기지 않는다
    }
  }
  return results;
}

/**
 * 채운 서식을 .hwpx로 저장(시험 출력·한글 파일로 저장). 저장 위치는 사용자가 고른다.
 * 한글이 있는 Windows에서 메일머지 필드 서식이면 한글이 원본을 열어 직접 바꾸고 저장한다(파일이 깨질 일이 없다).
 * 그 밖(한글이 없는 PC, 필드 없는 서식)에는 서식 엔진이 직접 채운다.
 */
async function saveFilled(library, entry, input, outPath, { asks = {}, today = new Date(), hancom = runHancom, platform = process.platform } = {}) {
  const source = await ensureHwpx(library, entry, { hancom });
  const buffer = fs.readFileSync(source);
  if (platform === 'win32') {
    const replacements = replacementsFor(entry, buffer, input, { asks, today });
    if (replacements) {
      const r = await hancom('printReplace', { file: library.sourcePath(entry), replacements, output: outPath });
      if (!r.ok) throw new Error(`한글로 서식을 채워 저장하지 못했습니다: ${r.error || '알 수 없는 오류'}`);
      return { mode: 'hancom', filled: replacements.length };
    }
  }
  const filled = fillForm(entry, buffer, input, { asks, today });
  fs.writeFileSync(outPath, filled.buffer);
  return filled;
}

module.exports = { fillForm, replacementsFor, ensureHwpx, printForms, saveFilled, cleanupTemp, valueFor, TEMP_PREFIX };
