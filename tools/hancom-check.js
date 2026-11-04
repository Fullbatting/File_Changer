// tools/hancom-check.js
// 한글 연결 점검: 이 PC에서 한글 자동 조작이 되는지, 서식이 제대로 인식·기입되는지 확인한다.
// 실행: tools\한글연결점검.bat (서식 .hwp/.hwpx 파일을 bat 위로 끌어다 놓으면 그 서식도 시험)
//   node tools/hancom-check.js [서식 파일 ...] [--print]
// 결과: 프로그램 폴더의 "한글연결점검_결과" 폴더에 보고서(점검보고서.txt)와 시험 파일(.hwpx·.pdf).
// 시험에 쓰는 인적사항은 모두 가짜 예시 값이다. 메일머지 필드({{이름}} 등)가 있는 서식은 서류 출력과 같은 방법으로,
// 필드가 없는 서식은 자동 인식으로 채운다. .hwp를 끌어다 놓으면 한글로 .hwpx로 바꾼 뒤 필드가 그대로인지도 본다.

const fs = require('fs');
const path = require('path');
const { runHancom } = require('../hancom.js');
const { modelFromHwpx } = require('../form-model.js');
const { recognize, summarize } = require('../form-recognize.js');
const { resolveValues } = require('../field-dictionary.js');
const { fillHwpx } = require('../form-fill.js');
const { fillForm } = require('../print-job.js');
const { inspectFormFile } = require('../form-library.js');
const { unknownFields, SAMPLE_INPUT } = require('../merge-fields.js');
const { readHwpFields } = require('../hwp-reader.js');



const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(ROOT, '한글연결점검_결과');
const args = process.argv.slice(2);
const doPrint = args.includes('--print');
const files = args.filter((a) => !a.startsWith('--'));

const lines = [];
const say = (s = '') => {
  lines.push(s);
  console.log(s);
};
const stepLine = (r) => (r.steps || []).map((s) => `${s.ok ? '○' : '×'} ${s.name}${s.detail ? ` (${s.detail})` : ''}`).join(' → ');

const SAMPLE = {
  values: {
    'patient.name': '홍길동',
    'patient.rrn': '900101-1234567',
    'patient.addr': '[예시 주소] ○○시 ○○로 00',
    'patient.phone': '000-000-0000',
    'patient.mobile': '010-0000-0000',
    'patient.email': 'example@example.com',
    'patient.household': '3',
    'patient.bank': '[예시은행]',
    'patient.account': '000-00-000000',
    'applicant.name': '김보호',
    'applicant.rrn': '600101-2345678',
    'applicant.relation': '배우자',
    'applicant.phone': '010-0000-0000',
    'diagnosis.name': '[예시 상병명]',
    'diagnosis.code': 'C00.0',
    'diagnosis.firstDate': '2026. 1. 2.',
  },
  org: { name: '[예시 병원]', phone: '000-000-0000', addr: '[예시 소재지]' },
  applicantMode: 'proxy',
};

async function timed(label, fn) {
  const t0 = Date.now();
  const r = await fn();
  r.ms = Date.now() - t0;
  say(`${r.ok ? '[성공]' : '[실패]'} ${label} — ${(r.ms / 1000).toFixed(1)}초${r.error ? ` — ${r.error}` : ''}`);
  if (r.steps && r.steps.length) say(`        ${stepLine(r)}`);
  return r;
}

async function checkForm(file) {
  const base = path.basename(file, path.extname(file));
  say('');
  say(`■ 서식 시험: ${path.basename(file)}`);
  let hwpx = file;
  if (path.extname(file).toLowerCase() === '.hwp') {
    hwpx = path.join(OUT, `${base}.hwpx`);
    const c = await timed('.hwp → .hwpx 변환(한글)', () => runHancom('convert', { input: file, output: hwpx }));
    if (!c.ok) return;
  }
  let model;
  let info;
  try {
    model = modelFromHwpx(fs.readFileSync(hwpx));
    info = inspectFormFile(hwpx);
  } catch (err) {
    say(`[실패] .hwpx 읽기 — ${err.message}`);
    return;
  }
  if (info.fields.length > 0) {
    // 메일머지 필드 서식: 서류 출력과 같은 방법(print-job.js)으로 채운다
    const unknown = unknownFields(info.fields);
    say(`[성공] 메일머지 필드 ${info.fields.length}개 · ${info.fields.reduce((n, f) => n + f.count, 0)}곳 — 서류 ${info.documents.length}종${unknown.length ? ` · 사전에 없는 필드: ${unknown.join(', ')}` : ''}`);
    info.documents.forEach((d) => say(`        ${d.mark} ${d.title}`.trimEnd()));
    if (path.extname(file).toLowerCase() === '.hwp') {
      const before = readHwpFields(fs.readFileSync(file));
      const same = JSON.stringify(before.fields) === JSON.stringify(info.fields);
      say(`${same ? '[성공]' : '[확인 필요]'} .hwp와 바꾼 .hwpx의 필드가 같은지 — ${same ? '같음' : `.hwp ${before.fields.length}개 / .hwpx ${info.fields.length}개`}`);
    }
    const filled = fillForm({ mapping: {} }, fs.readFileSync(hwpx), SAMPLE_INPUT, { asks: Object.fromEntries(unknown.map((n) => [n, `[${n}]`])) });
    const leftover = (modelFromHwpx(filled.buffer).paragraphs.map((p) => p.text).join('\n').match(/\{\{[^{}]+\}\}/g) || []).length;
    const filledPath = path.join(OUT, `${base}_채움시험.hwpx`);
    fs.writeFileSync(filledPath, filled.buffer);
    say(`${leftover === 0 ? '[성공]' : '[확인 필요]'} 예시 값으로 필드 ${filled.filled}곳 채우기${leftover ? ` — 안 채워진 {{필드}} ${leftover}곳` : ''} → ${path.basename(filledPath)}`);
    await timed('채운 서식을 한글로 열어 PDF로 저장(모양 확인용)', () => runHancom('convert', { input: filledPath, output: path.join(OUT, `${base}_채움시험.pdf`) }));
    return;
  }
  const spots = recognize(model);
  const s = summarize(spots);
  say(`[성공] 기입 위치 자동 인식 — ${s.total}곳 (${Object.entries(s.byKind).map(([k, v]) => `${k} ${v}`).join(', ')}) · 확인 필요 ${s.check} · 맞는 항목 없음 ${s.none}`);
  spots.forEach((x) => say(`        ${String(x.no).padStart(2)}. [${x.kind}] ${x.where} ${x.group ? `${x.group} › ` : ''}${x.label} → ${x.field || (x.options ? x.options.map((o) => o.text).join(' / ') : x.ask ? '(출력 때 입력)' : '(비워 둠)')}${x.confidence === 'check' ? ' ※확인 필요' : ''}`));
  const choices = {};
  spots.filter((x) => x.kind === 'choice').forEach((x) => (choices[x.id] = [0]));
  const asks = {};
  spots.filter((x) => x.ask).forEach((x) => (asks[x.id] = '[예시]'));
  const filled = fillHwpx(fs.readFileSync(hwpx), spots, { values: resolveValues(SAMPLE), choices, asks });
  const filledPath = path.join(OUT, `${base}_채움시험.hwpx`);
  fs.writeFileSync(filledPath, filled.buffer);
  say(`[성공] 예시 값으로 채우기 — ${filled.plan.length}곳 → ${path.basename(filledPath)}`);
  await timed('채운 서식을 한글로 열어 PDF로 저장(모양 확인용)', () => runHancom('convert', { input: filledPath, output: path.join(OUT, `${base}_채움시험.pdf`) }));
}

async function main() {
  fs.mkdirSync(OUT, { recursive: true });
  say('서식 자동 기입 — 한글 연결 점검');
  say(`시각: ${new Date().toLocaleString('ko-KR')} · Windows: ${process.platform === 'win32' ? '예' : `아니오(${process.platform})`} · Node ${process.version}`);
  say('');
  const probe = await timed('한글 연결', () => runHancom('probe'));
  if (probe.ok) {
    say(`        한글 버전: ${probe.version || '(알 수 없음)'}`);
    say(`        보안 모듈: ${probe.securityModule ? '등록됨 — 파일 접근 확인 창 없이 진행' : '등록 안 됨 — 파일을 열 때마다 "접근 허용" 확인 창이 뜰 수 있음'}`);
    if (probe.registeredModules && probe.registeredModules.length) say(`        등록된 모듈: ${[].concat(probe.registeredModules).join(', ')}`);
  }
  if (probe.ok) {
    const self = await timed('새 문서를 .hwpx·.pdf로 저장하고 다시 열기', () => runHancom('selftest', { hwpxOut: path.join(OUT, '연결시험.hwpx'), pdfOut: path.join(OUT, '연결시험.pdf') }));
    for (const f of files) await checkForm(path.resolve(f));
    if (doPrint && self.ok) await timed('기본 프린터로 연결시험 문서 1부 인쇄', () => runHancom('print', { file: path.join(OUT, '연결시험.hwpx'), copies: 1 }));
    else if (!doPrint) say('\n(인쇄 시험은 하지 않았습니다. 하려면 bat 파일 대신 명령 창에서 --print 를 붙여 실행하세요.)');
  } else if (files.length) {
    say('\n한글 연결이 안 되어 서식 시험은 건너뛰었습니다.');
  }
  say('');
  say(`결과 폴더: ${OUT}`);
  say('이 폴더의 점검보고서.txt와 .hwpx·.pdf 파일을 보내 주세요(시험 값은 모두 가짜 예시입니다).');
  fs.writeFileSync(path.join(OUT, '점검보고서.txt'), `﻿${lines.join('\r\n')}\r\n`, 'utf8');
}

main().catch((err) => {
  say(`[오류] ${err && err.stack ? err.stack : err}`);
  try {
    fs.writeFileSync(path.join(OUT, '점검보고서.txt'), `﻿${lines.join('\r\n')}\r\n`, 'utf8');
  } catch (e) {
    /* 보고서를 못 써도 화면에는 나왔다 */
  }
  process.exit(1);
});
