// form-fill.js
// 기입 위치(form-recognize.js)마다 넣을 글자를 정하고, .hwpx 원문 XML의 그 자리만 바꿔 새 파일을 만든다.
// - 빈 칸: 값을 그대로 넣는다(칸 서식·글꼴은 그 칸의 첫 런을 따른다). "-"만 찍힌 주민등록번호 칸도 통째로 바꾼다.
// - 한 칸에 머리글과 값을 쓰는 공통서식: 머리글 아래 빈 줄에, 빈 줄이 없으면 머리글 뒤에 두 칸 띄우고 이어 쓴다.
// - "라벨:      " 빈자리·서명 줄: 빈칸 너비를 크게 바꾸지 않도록 값 길이만큼 공백을 줄인다(한글 2칸, 그 밖 1칸으로 셈).
// - 날짜 틀: 년·월·일 앞 공백 자리에 숫자를 넣는다.
// - 선택 표시: 고른 보기의 "[   ]"를 "[ V ]"로(괄호 안 글자 수는 그대로).
// 화면에 기대지 않아 node로 단위 테스트한다.

const PizZip = require('pizzip');
const { applyEdits, encodeText } = require('./hwpx-xml.js');
const { modelFromHwpx } = require('./form-model.js');

const CHECK_MARK = 'V';

/** 화면 너비 추정: 한글·한자·전각 2칸, 나머지 1칸 */
function displayWidth(s) {
  let w = 0;
  for (const ch of String(s)) {
    const c = ch.codePointAt(0);
    w += (c >= 0x1100 && c <= 0x11ff) || (c >= 0x2e80 && c <= 0xa4cf) || (c >= 0xac00 && c <= 0xd7a3) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0xff00 && c <= 0xff60) ? 2 : 1;
  }
  return w;
}

/** n칸짜리 공백 자리에 값을 넣은 글자: " 값" + 남은 공백(최소 1칸) */
function padInto(n, value) {
  return ` ${value}${' '.repeat(Math.max(1, n - 1 - displayWidth(value)))}`;
}

/** 괄호 안 L칸에 체크 표시: L=3이면 " V " → "[ V ]" */
function checkInto(L) {
  if (L <= 1) return CHECK_MARK;
  const left = Math.floor((L - 1) / 2);
  return ' '.repeat(left) + CHECK_MARK + ' '.repeat(L - 1 - left);
}

/** 기입 위치의 연결: 등록 화면에서 정한 mapping이 있으면 그것, 없으면 인식 결과 그대로 */
function mappingOf(spot) {
  if (spot.mapping) return spot.mapping;
  if (spot.kind === 'choice') return { type: 'choice', defaults: [] };
  if (spot.ask) return { type: 'ask' };
  if (spot.field) return { type: 'field', key: spot.field };
  return { type: 'blank' };
}

/**
 * 채울 글자 계획: [{ spotId, paragraphId, start, end, text }] (문단 글자 기준 위치).
 * answers: { values(resolveValues 결과), choices: { [spotId]: [보기 번호] }, asks: { [spotId]: 글자 } }
 */
function planFill(model, spots, answers = {}) {
  const values = answers.values || {};
  const plan = [];
  const chosen = new Map(); // spotId -> Set(보기 번호)
  spots
    .filter((s) => s.kind === 'choice')
    .forEach((s) => {
      const m = mappingOf(s);
      const picked = answers.choices && answers.choices[s.id] !== undefined ? answers.choices[s.id] : m.defaults || [];
      chosen.set(s.id, new Set([].concat(picked).map(Number)));
    });
  const active = (s) => !s.dependsOn || (chosen.get(s.dependsOn.spotId) || new Set()).has(s.dependsOn.option);
  const push = (spot, paragraphId, start, end, text) => plan.push({ spotId: spot.id, paragraphId, start, end, text });

  spots.forEach((spot) => {
    if (!active(spot)) return;
    const m = mappingOf(spot);
    if (spot.kind === 'choice') {
      (chosen.get(spot.id) || new Set()).forEach((i) => {
        const o = spot.options[i];
        if (o) push(spot, o.paragraphId, o.range.start, o.range.end, checkInto(o.range.end - o.range.start));
      });
      return;
    }
    if (m.type === 'blank') return;
    let value = '';
    if (m.type === 'fixed') value = m.text || '';
    else if (m.type === 'ask') value = (answers.asks && answers.asks[spot.id]) || '';
    else if (m.type === 'field') value = values[m.key] || '';
    value = String(value).trim();
    if (spot.kind === 'date' && m.type === 'field' && m.key === 'auto.today') {
      const t = values.__today;
      if (!t) return;
      [['year', t.y], ['month', t.m], ['day', t.d]].forEach(([part, num]) => {
        const r = spot.parts[part].range;
        const n = r.end - r.start;
        const s = String(num);
        push(spot, spot.parts[part].paragraphId, r.start, r.end, ' '.repeat(Math.max(0, n - s.length)) + s);
      });
      return;
    }
    if (!value) return;
    const p = model.paragraphById.get(spot.paragraphId);
    const text = p ? p.text : '';
    const { start, end } = spot.range;
    if (spot.kind === 'incell' && spot.append) {
      push(spot, spot.paragraphId, text.length, text.length, `  ${value}`);
    } else if (spot.kind === 'cell' || spot.kind === 'below' || spot.kind === 'incell') {
      push(spot, spot.paragraphId, 0, text.length, value); // 빈 칸(공백뿐일 수도 있음)을 통째로
    } else if (spot.kind === 'unit') {
      push(spot, spot.paragraphId, start, start, `${value} `);
    } else if (spot.kind === 'inline' || spot.kind === 'sign') {
      const n = end - start;
      // 서명 칸이 따로 있는 줄은 "(서명 또는 인)" 앞에 이름을 넣는다
      push(spot, spot.paragraphId, start, end, n === 0 ? (spot.kind === 'sign' ? `${value}  ` : ` ${value}`) : padInto(n, value));
    } else {
      push(spot, spot.paragraphId, start, end, value); // placeholder
    }
  });
  return plan;
}

/** 계획을 적용한 문단 글자(미리보기·검사용): Map(paragraphId -> 글자) */
function previewTexts(model, plan) {
  const byP = new Map();
  plan.forEach((e) => {
    if (!byP.has(e.paragraphId)) byP.set(e.paragraphId, []);
    byP.get(e.paragraphId).push(e);
  });
  const out = new Map();
  byP.forEach((edits, pid) => {
    let text = model.paragraphById.get(pid).text;
    edits
      .slice()
      .sort((a, b) => b.start - a.start)
      .forEach((e) => {
        text = text.slice(0, e.start) + e.text + text.slice(e.end);
      });
    out.set(pid, text);
  });
  return out;
}

const encodeValue = (s) => encodeText(s).replace(/\r\n|\r|\n/g, '<hp:lineBreak/>');

/** 문단 하나에 대한 글자 편집들을 원문 XML 편집으로 바꾼다 */
function paragraphXmlEdits(xml, p, edits) {
  const textPieces = p.pieces.filter((x) => x.kind === 'text');
  const prefix = p.el.name.includes(':') ? p.el.name.slice(0, p.el.name.indexOf(':') + 1) : '';
  if (textPieces.length === 0) {
    // 글자가 하나도 없는 문단(빈 칸): 첫 런 안에 <hp:t>를 만든다
    const value = edits.map((e) => e.text).join('');
    const run = p.runs[0];
    if (!run) return [{ start: p.el.openEnd, end: p.el.openEnd, text: `<${prefix}run charPrIDRef="0"><${prefix}t>${encodeValue(value)}</${prefix}t></${prefix}run>` }];
    const emptyT = run.children.find((c) => c.type === 'el' && c.local === 't');
    if (emptyT) {
      if (emptyT.selfClosing) return [{ start: emptyT.start, end: emptyT.end, text: `<${emptyT.name}${emptyT.attrs.replace(/\/\s*$/, '')}>${encodeValue(value)}</${emptyT.name}>` }];
      return [{ start: emptyT.openEnd, end: emptyT.openEnd, text: encodeValue(value) }];
    }
    if (run.selfClosing) return [{ start: run.start, end: run.end, text: `<${run.name}${run.attrs.replace(/\/\s*$/, '')}><${prefix}t>${encodeValue(value)}</${prefix}t></${run.name}>` }];
    return [{ start: run.openEnd, end: run.openEnd, text: `<${prefix}t>${encodeValue(value)}</${prefix}t>` }];
  }
  // 글자 조각마다 자기 구간의 편집을 모은다. 여러 조각에 걸친 편집은 첫 조각에 새 글자를, 나머지는 지우기만.
  const local = new Map(); // piece -> [{ s, e, text }]
  const addLocal = (piece, s, e, text) => {
    if (!local.has(piece)) local.set(piece, []);
    local.get(piece).push({ s: s - piece.plainStart, e: e - piece.plainStart, text });
  };
  edits.forEach((ed) => {
    if (ed.start === ed.end) {
      const host = textPieces.find((x) => x.plainStart <= ed.start && ed.start < x.plainEnd) || textPieces.find((x) => x.plainEnd === ed.start) || textPieces[0];
      const at = Math.min(Math.max(ed.start, host.plainStart), host.plainEnd);
      addLocal(host, at, at, ed.text);
      return;
    }
    let first = true;
    textPieces.forEach((x) => {
      const s = Math.max(ed.start, x.plainStart);
      const e = Math.min(ed.end, x.plainEnd);
      if (s >= e) return;
      addLocal(x, s, e, first ? ed.text : '');
      first = false;
    });
    if (first) {
      // 구간이 글자 조각이 아닌 곳(줄바꿈 등)만 덮으면 바로 앞 조각 끝에 넣는다
      const host = textPieces.filter((x) => x.plainEnd <= ed.start).pop() || textPieces[0];
      addLocal(host, host.plainEnd, host.plainEnd, ed.text);
    }
  });
  const out = [];
  local.forEach((list, piece) => {
    let s = piece.decoded;
    list
      .slice()
      .sort((a, b) => b.s - a.s || b.e - a.e)
      .forEach((x) => {
        s = s.slice(0, x.s) + '\u0000' + encodeValue(x.text).replace(/</g, '\u0001') + '\u0002' + s.slice(x.e);
      });
    // 원래 글자는 다시 이스케이프하고, 새로 넣은 부분(줄바꿈 요소 포함)은 그대로 둔다
    const raw = s
      .split(/(\u0000[^\u0002]*\u0002)/)
      .map((part) => (part.startsWith('\u0000') ? part.slice(1, -1).replace(/\u0001/g, '<') : encodeText(part)))
      .join('');
    out.push({ start: piece.node.start, end: piece.node.end, text: raw });
  });
  return out;
}

/**
 * 고친 문단의 줄 배치 정보(<hp:linesegarray>)를 지우는 편집. 한글은 문단마다 "몇 번째 글자에서 줄이 바뀌는지"를
 * 저장해 두는데, 글자를 바꾸면 이 값이 맞지 않게 된다(좁은 칸의 "{{1}}"이 세 줄로 저장돼 있다가 한 글자가 되는 등).
 * 맞지 않는 값이 남으면 한글이 파일이 손상됐다며 열지 않는다. 지우면 한글이 열 때 줄을 다시 나눈다.
 */
function staleLayoutEdits(p) {
  return p.el.children.filter((c) => c.type === 'el' && c.local === 'linesegarray').map((c) => ({ start: c.start, end: c.end, text: '' }));
}

/**
 * zip 안 항목마다 원래 압축 방식(0 = 저장만, 8 = 압축). 중앙 디렉터리를 직접 읽는다.
 * 한글이 저장한 .hwpx는 mimetype·version.xml·미리보기 그림을 압축하지 않고 두는데, 이것을 압축해 다시 쓰면
 * 한글이 파일을 열지 못한다. 그래서 다시 쓸 때 항목마다 원래 방식을 지킨다.
 */
function zipMethods(buffer) {
  const methods = new Map();
  const min = Math.max(0, buffer.length - 65557);
  let eocd = -1;
  for (let i = buffer.length - 22; i >= min; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return methods;
  const count = buffer.readUInt16LE(eocd + 10);
  let p = buffer.readUInt32LE(eocd + 16);
  for (let n = 0; n < count && p + 46 <= buffer.length && buffer.readUInt32LE(p) === 0x02014b50; n++) {
    const method = buffer.readUInt16LE(p + 10);
    const nameLen = buffer.readUInt16LE(p + 28);
    const extraLen = buffer.readUInt16LE(p + 30);
    const commentLen = buffer.readUInt16LE(p + 32);
    methods.set(buffer.slice(p + 46, p + 46 + nameLen).toString('utf8'), method);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return methods;
}

/**
 * 글자 계획(planFill이나 메일머지가 만든 [{ spotId, paragraphId, start, end, text }])을 원문 XML에 적용한 새 .hwpx 버퍼.
 * model은 이 buffer로 만든 모델이어야 한다(바꾸지 않으므로 여러 번 다시 써도 된다 — 메일머지 일괄 생성).
 */
function writePlan(buffer, model, plan) {
  const zip = new PizZip(buffer);
  const bySection = new Map();
  plan.forEach((e) => {
    const p = model.paragraphById.get(e.paragraphId);
    if (!p) throw new Error(`서식에서 기입 위치를 찾을 수 없습니다(${e.spotId}). 서식 파일이 바뀌었으면 다시 등록하세요.`);
    if (!bySection.has(p.section)) bySection.set(p.section, new Map());
    const m = bySection.get(p.section);
    if (!m.has(p)) m.set(p, []);
    m.get(p).push(e);
  });
  bySection.forEach((paragraphs, si) => {
    const sec = model.sections[si];
    const edits = [];
    paragraphs.forEach((list, p) => edits.push(...paragraphXmlEdits(sec.xml, p, list), ...staleLayoutEdits(p)));
    zip.file(sec.name, applyEdits(sec.xml, edits));
  });
  // 항목마다 원래 압축 방식을 지킨다(mimetype·version.xml·미리보기 그림은 한글이 압축하지 않고 둔다).
  // 이름을 그대로 다시 넣으므로 순서(mimetype이 맨 앞)도 그대로다
  const methods = zipMethods(buffer);
  if (!methods.has('mimetype')) methods.set('mimetype', 0);
  methods.forEach((method, name) => {
    const entry = zip.file(name);
    if (entry && method === 0) zip.file(name, entry.asUint8Array(), { binary: true, compression: 'STORE' });
  });
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/**
 * .hwpx(버퍼)의 기입 위치를 채운 새 .hwpx(버퍼). spots는 이 파일에서 인식한(또는 저장해 둔) 기입 위치.
 * 반환: { buffer, plan }
 */
function fillHwpx(buffer, spots, answers = {}) {
  const model = modelFromHwpx(buffer);
  const plan = planFill(model, spots, answers);
  return { buffer: writePlan(buffer, model, plan), plan };
}

module.exports = { fillHwpx, writePlan, zipMethods, planFill, previewTexts, mappingOf, displayWidth, padInto, checkInto, CHECK_MARK };
