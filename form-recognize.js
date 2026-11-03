// form-recognize.js
// 서식 모델(form-model.js)에서 "내용을 써 넣을 자리(기입 위치)"를 규칙으로 찾는다. AI를 쓰지 않는다.
// 찾는 자리:
//  - cell   : 짧은 머리글 칸 바로 오른쪽의 빈 칸 (예: [성명][    ])
//  - below  : 머리글만 있는 줄 바로 아래의 빈 칸 (예: 첫 줄 [성명][관계], 둘째 줄 빈 칸)
//  - incell : 머리글과 값을 한 칸에 쓰는 공통서식 (예: [성  명        ] 넓은 칸, 또는 머리글 줄 아래 빈 줄)
//  - unit   : 단위만 든 칸의 단위 앞 (예: [가구원수][    명])
//  - inline : 칸·문단 안 "라벨: (빈칸)" (예: 금융기관:      예금주:      계좌번호:)
//  - sign   : 서명 줄 "신청인        (서명 또는 인)"
//  - date   : 날짜 틀 "    년    월    일"
//  - choice : 선택 표시 "[   ] 동의   [   ] 동의하지않음" — 괄호 안 하위 선택은 따로 묶는다
//  - placeholder : {{변수}}
// 같은 머리글이 여러 번 나오면(성명·전화번호) 표 왼쪽 묶음 머리글(등록 대상자 · 신청인 · 진단정보)로 누구 것인지 정한다.

const { cellAt } = require('./form-model.js');
const { matchLabel, groupKind, normalizeLabel } = require('./field-dictionary.js');

// 빈 칸: 공백뿐이거나, 주민등록번호 칸처럼 "-"만 미리 찍혀 있는 칸
const BLANK_CELL = (t) => /^[-－]?$/.test(String(t || '').replace(/\s/g, ''));
// 문서 제목(○○신청서·동의서·위임장 …)은 머리글이 아니다
const TITLE = /(신청서|동의서|확인서|신고서|요청서|위임장|통보서|청구서|보고서|조사서|추천서|지정서)$/;
const UNIT_ONLY = /^\s*(명|원|세|개월|년|월|일|회|%|kg|cm|매|부)\s*$/;
const SIGN_WORDS = /(신청인|대리신청인|동의자|민원인|신청자|제출인|위임인|보호자|확인자|신고자|신고인|요청인)/;
// 시청이 쓰는 칸(접수·처리·결재)과 받는 곳: 기입 위치로 보지 않는다
const OFFICE_LABEL = /^(접수번호|접수일자?|처리기간|처리절차|처리기관|발급번호|결재|검토|접수|담당자.*|.*귀하)$/;
const ARROW_CELL = /^\s*[➡→⇨▶⇒]+\s*$/;
const INLINE_BLANK = /[:：]\s*($|[)）]|\s{2,})/;

/** 머리글로 볼 만한 칸: 짧고, 선택 표시·빈칸 틀이 없고, 숫자뿐이 아님 */
function isLabelText(text) {
  const t = String(text || '').trim();
  if (!t || t.length > 30) return false;
  if (/\[\s*\]/.test(t) || /[:：]\s*$/.test(t) || /[:：]\s{2,}/.test(t)) return false;
  if (/^[\d\s.,\-/()]+$/.test(t)) return false;
  if (UNIT_ONLY.test(t) || ARROW_CELL.test(t)) return false;
  if (TITLE.test(t.replace(/[\s()（）]/g, ''))) return false;
  if (/(년\s+월\s+일)/.test(t)) return false;
  return true;
}

const oneLine = (s) => String(s || '').replace(/\s+/g, ' ').trim();

/**
 * 칸의 묶음 머리글: 같은 행을 덮는 맨 왼쪽(0열) 칸이 이 칸과 다르고, 여러 행을 묶거나(rowSpan>1)
 * 그 오른쪽에 다시 머리글이 오는 구조일 때. 반환: { text, kind } 또는 null
 */
function groupOf(table, cell) {
  if (cell.col === 0) return null;
  const left = cellAt(table, cell.row, 0);
  if (!left || left === cell || !isLabelText(left.text)) return null;
  // 머리글 줄의 첫 칸([관계][성명][주민등록번호] 아래가 빈 줄)은 묶음이 아니라 열 머리글
  const under = left.rowSpan === 1 ? cellAt(table, left.row + 1, left.col) : null;
  if (under && under.colSpan === left.colSpan && BLANK_CELL(under.text)) return null;
  const next = cellAt(table, cell.row, left.col + left.colSpan);
  const groupsRows = left.rowSpan > 1 || (next && isLabelText(next.text));
  if (!groupsRows) return null;
  return { text: oneLine(left.text), kind: groupKind(left.text) };
}

/** 사람이 읽는 위치 설명 */
const whereOf = (table, cell) => `표${table.id.slice(1)} ${cell.row + 1}행 ${cell.col + 1}열`;

function recognize(model) {
  const spots = [];
  const used = new Map(); // paragraphId -> [[start,end]] 이미 쓴 글자 구간(겹치는 자리 방지)
  const take = (pid, start, end) => {
    const list = used.get(pid) || [];
    if (list.some(([a, b]) => (start < b && a < end) || (start === end && a <= start && start < b && a !== b))) return false;
    list.push([start, end]);
    used.set(pid, list);
    return true;
  };
  const add = (spot) => {
    spots.push({ id: `${spot.kind}:${spot.paragraphId || (spot.parts && spot.parts.year.paragraphId)}:${spot.range ? spot.range.start : spot.parts ? spot.parts.year.range.start : 0}`, ...spot });
    return spots[spots.length - 1];
  };
  const mapLabel = (label, group, fallbackWhy) => {
    if (group && group.kind === 'office') return { field: null, confidence: 'none', why: '시청(보장기관) 쪽 칸입니다 — 서식 등록 때 고정값으로 한 번 넣어 두세요' };
    if (group && group.kind === 'work') return { field: null, confidence: 'none', why: '직장 정보 칸입니다(환자 인적사항이 아님)' };
    const m = matchLabel(label, group ? group.kind : null);
    if (!m) return { field: null, confidence: 'none', why: fallbackWhy || '맞는 입력 항목이 없습니다' };
    return m;
  };
  const paragraphOf = (id) => model.paragraphById.get(id);
  // 표·그림이 들어앉은 칸: 글자가 없어 빈 칸처럼 보여도 쓰지 않는다
  const blockedCell = (cell) => cell.paragraphIds.some((id) => paragraphOf(id).hasObject);

  // 문서에 신청인 묶음이 있는지(묶음 없는 "성명"이 누구 것인지 모호한지 판단)
  const hasApplicantGroup = model.tables.some((t) => t.cells.some((c) => c.col === 0 && groupKind(c.text) === 'applicant'));

  /* ---------- 1. 표 칸 ---------- */
  model.tables.forEach((table) => {
    // 처리 절차 그림(➡이 든 줄과 그 아래 줄)은 건너뛴다
    const flowRows = new Set();
    table.cells.filter((c) => ARROW_CELL.test(c.text)).forEach((c) => {
      flowRows.add(c.row);
      flowRows.add(c.row + c.rowSpan);
    });
    table.cells.forEach((cell) => {
      if (flowRows.has(cell.row)) return;
      if (table.colCnt > 1 && cell.colSpan === table.colCnt) return; // 표 너비 전체 칸: 제목·안내 문구
      const firstP = cell.paragraphIds.map(paragraphOf).find((p) => p.text.trim());
      if (!firstP) return;
      if (OFFICE_LABEL.test(normalizeLabel(firstP.text))) return;
      if (!isLabelText(cell.text)) {
        incell(table, cell);
        return;
      }
      const group = groupOf(table, cell);
      // 묶음 머리글 칸 자체(예: "등록 대상자")는 자리가 아니다
      if (cell.col === 0) {
        const next = cellAt(table, cell.row, cell.colSpan);
        if (next && isLabelText(next.text) && cell.rowSpan > 1) return;
      }
      const right = cellAt(table, cell.row, cell.col + cell.colSpan);
      const label = oneLine(cell.text);
      const describe = (m, extraWhy) => {
        let { confidence, why } = m;
        if (m.field && m.field.startsWith('patient.') && !group && hasApplicantGroup && /^(성명|이름|주민등록번호|전화번호|연락처)$/.test(normalizeLabel(label))) {
          confidence = 'check';
          why = '묶음 머리글이 없어 환자 것인지 신청인 것인지 정해야 합니다';
        }
        return { confidence, why: why || extraWhy || '' };
      };
      if (right && right !== cell && BLANK_CELL(right.text) && right.paragraphIds.length > 0 && !blockedCell(right)) {
        const m = mapLabel(label, group);
        add({ kind: 'cell', paragraphId: right.paragraphIds[0], range: { start: 0, end: 0 }, label, group: group && group.text, where: whereOf(table, right), field: m.field, ...describe(m), reason: `표 칸 · 왼쪽 머리글 “${label}”${group ? ` (${group.text})` : ''}` });
        return;
      }
      if (right && right !== cell && UNIT_ONLY.test(right.text)) {
        const p = right.paragraphIds.map(paragraphOf).find((x) => UNIT_ONLY.test(x.text));
        if (p) {
          const at = p.text.search(/\S/);
          const m = mapLabel(label, group);
          add({ kind: 'unit', paragraphId: p.id, range: { start: at, end: at }, unit: p.text.trim(), label, group: group && group.text, where: whereOf(table, right), field: m.field, ...describe(m), reason: `표 칸 · 왼쪽 머리글 “${label}” · 단위 “${p.text.trim()}” 앞` });
          return;
        }
      }
      // 머리글 줄 아래의 빈 칸: 이 칸 오른쪽도 머리글이고(머리글만 있는 줄), 바로 아래 칸이 비어 있음
      const below = cellAt(table, cell.row + cell.rowSpan, cell.col);
      // 머리글 줄: 칸마다 글자가 있고 빈자리 틀이 없다(“금융정보 등의 제공을 동의함 (서명)”처럼 긴 머리글도)
      const headerText = (c) => isLabelText(c.text) || (c.text.trim() !== '' && !/\[\s*\]/.test(c.text) && !INLINE_BLANK.test(c.text));
      const rowIsHeader = table.cells.filter((c) => c.row === cell.row).every((c) => c.col === 0 || headerText(c));
      const separator = below && table.colCnt > 1 && below.colSpan === table.colCnt; // 표 너비 전체 빈 줄은 칸 사이 띄움
      if (rowIsHeader && below && below !== cell && !separator && below.col === cell.col && BLANK_CELL(below.text) && !blockedCell(below) && below.paragraphIds.length > 0 && (!right || right === cell || headerText(right))) {
        const m = mapLabel(label, group);
        add({ kind: 'below', paragraphId: below.paragraphIds[0], range: { start: 0, end: 0 }, label, group: group && group.text, where: whereOf(table, below), field: m.field, ...describe(m), reason: `표 칸 · 위 머리글 “${label}”` });
        return;
      }
      // 서명 칸이 따로 있는 줄: [요청인][(서명 또는 인)]
      if (right && right !== cell && /^\s*[(（]\s*(서명|인|날인)/.test(right.text) && SIGN_WORDS.test(label)) {
        add({ kind: 'sign', paragraphId: right.paragraphIds[0], range: { start: 0, end: 0 }, label, where: whereOf(table, right), field: 'auto.signer', confidence: 'sure', why: '', reason: `서명 칸 · “${label}” 오른쪽 “(서명 또는 인)” 앞` });
        return;
      }
      incell(table, cell);
    });
  });

  /**
   * 머리글과 값을 한 칸에 쓰는 공통서식 칸. 오른쪽·아래 빈 칸이 없을 때만 부른다.
   * - 머리글 문단 뒤에 빈 문단이 있으면 그 빈 문단에 ([성명 / (빈 줄)])
   * - 없으면 머리글 뒤에 이어 쓴다([성  명        ] 넓은 칸) — 맞는 입력 항목이 있을 때만
   */
  function incell(table, cell) {
    if (cell.col === 0 && cell.rowSpan > 1) return; // 묶음 머리글
    const ps = cell.paragraphIds.map(paragraphOf);
    // 칸 전체가 머리글이면("동의자 / 성  명") 마지막 글자 줄까지가 머리글, 아니면 첫 글자 줄만
    const whole = isLabelText(cell.text);
    const textIdx = ps.map((p, i) => (p.text.trim() ? i : -1)).filter((i) => i >= 0);
    const at = whole ? textIdx[textIdx.length - 1] : textIdx[0];
    const labelP = ps[at];
    const label = oneLine(whole ? cell.text : labelP.text);
    if (!isLabelText(label) || ps.some((p) => p.hasObject)) return;
    const group = groupOf(table, cell);
    if (cell.col === 0 && !group) {
      const next = cellAt(table, cell.row, cell.colSpan);
      if (next && isLabelText(next.text)) return; // 한 줄짜리 묶음 머리글("긴급의료지원 대상자")
    }
    const m = mapLabel(label, group);
    const rest = ps.slice(at + 1);
    const emptyP = rest.find((p) => BLANK_CELL(p.text) && !p.hasObject);
    const base = { label, group: group && group.text, where: whereOf(table, cell), field: m.field, confidence: m.confidence, why: m.why };
    if (emptyP) {
      add({ kind: 'incell', paragraphId: emptyP.id, range: { start: 0, end: 0 }, ...base, reason: `한 칸 · 머리글 “${label}” 아래 빈 줄` });
      return;
    }
    if (!m.field) return;
    // 칸 안에 "ㅇ 자택 :  ㅇ 휴대전화 :" 같은 빈자리가 따로 있으면 전화번호는 그쪽에 쓴다
    if (/(phone|mobile)$/.test(m.field) && rest.some((p) => INLINE_BLANK.test(p.text))) return;
    const end = labelP.text.length;
    if (!take(labelP.id, end, end)) return;
    add({ kind: 'incell', paragraphId: labelP.id, range: { start: end, end }, append: true, ...base, reason: `한 칸 · 머리글 “${label}” 뒤에 이어 씀` });
  }

  /* ---------- 2. 문단 안 패턴 ---------- */
  const cellGroupOf = (p) => {
    if (!p.cell) return null;
    const table = model.tableById.get(p.cell.tableId);
    const cell = cellAt(table, p.cell.row, p.cell.col);
    return cell ? groupOf(table, cell) : null;
  };
  const whereP = (p) => {
    if (!p.cell) return '본문 문단';
    return `표${p.cell.tableId.slice(1)} ${p.cell.row + 1}행 ${p.cell.col + 1}열`;
  };

  model.paragraphs.forEach((p) => {
    const text = p.text || '';
    if (!text.trim()) return;
    const group = cellGroupOf(p);
    let m;

    // {{변수}}
    const ph = /\{\{\s*([^{}]+?)\s*\}\}/g;
    while ((m = ph.exec(text))) {
      if (!take(p.id, m.index, m.index + m[0].length)) continue;
      const mm = mapLabel(m[1], group, `“${m[1]}”에 맞는 입력 항목이 없습니다`);
      add({ kind: 'placeholder', paragraphId: p.id, range: { start: m.index, end: m.index + m[0].length }, label: m[1], group: group && group.text, where: whereP(p), field: mm.field, confidence: mm.confidence, why: mm.why, reason: `서식에 적힌 {{${m[1]}}}` });
    }

    // 서명 줄: "신청인(대리신청인)：        (서명 또는 인)" · "민원인        (서명 또는 인)"
    const sign = /([가-힣()（）]{2,16}(?: [가-힣]{1,6})?)\s*[:：]?(\s{3,})[(（]\s*(서명|인|날인)/g;
    while ((m = sign.exec(text))) {
      const label = m[1].replace(/^[()（）]+/, '');
      if (!SIGN_WORDS.test(label)) continue;
      const start = m.index + m[0].indexOf(m[2]);
      const end = start + m[2].length;
      if (!take(p.id, start, end)) continue;
      const mm = matchLabel(label.replace(/[(（].*$/, ''), null) || { field: 'auto.signer', confidence: 'sure', why: '' };
      add({ kind: 'sign', paragraphId: p.id, range: { start, end }, label: oneLine(label), where: whereP(p), field: mm.field === 'auto.signer' ? 'auto.signer' : 'auto.signer', confidence: 'sure', why: '', reason: `서명 줄 · “${oneLine(label)} … (서명)”` });
    }

    // 날짜 틀: "      년      월      일" (숫자가 아직 없는 것만)
    const date = /(\s{2,})년(\s{2,})월(\s{2,})일/g;
    while ((m = date.exec(text))) {
      const yEnd = m.index + m[1].length;
      // "20    년"처럼 앞 두 자리가 미리 찍힌 틀은 그 숫자까지 연도 자리로
      const century = /(\d{2})$/.exec(text.slice(0, m.index));
      const yStart = century ? m.index - century[1].length : Math.max(m.index, yEnd - 8);
      const mStart = yEnd + 1;
      const mEnd = mStart + m[2].length;
      const dStart = mEnd + 1;
      const dEnd = dStart + m[3].length;
      if (!take(p.id, yStart, dEnd + 1)) continue;
      add({
        kind: 'date',
        paragraphId: p.id,
        range: { start: yStart, end: dEnd },
        parts: { year: { paragraphId: p.id, range: { start: yStart, end: yEnd } }, month: { paragraphId: p.id, range: { start: mStart, end: mEnd } }, day: { paragraphId: p.id, range: { start: dStart, end: dEnd } } },
        label: '년 월 일',
        where: whereP(p),
        field: 'auto.today',
        confidence: 'sure',
        why: '',
        reason: '날짜 틀 · “    년    월    일”',
      });
    }

    // 선택 표시는 아래 3단계에서 칸 단위로 묶는다. 여기서는 선택지 뒤 괄호 구간을 기억해 둔다.

    // 칸·문단 안 "라벨: (빈칸)"
    const inline = /([^\s:：[\]()（）,]{1,12}(?: [^\s:：[\]()（）,]{1,8})?)\s*[:：]( *)(?=$|[^ ])/g;
    while ((m = inline.exec(text))) {
      const label = m[1];
      const spaces = m[2];
      const after = text.slice(m.index + m[0].length);
      const next = after[0];
      const atEnd = after.trim() === '' || next === '\n';
      const closes = next === ',' || next === ')' || next === '）';
      if (!(spaces.length >= 2 || atEnd || closes)) continue;
      if (next === '[') continue;
      const start = m.index + m[0].length - spaces.length;
      const end = atEnd ? text.length : m.index + m[0].length;
      if (!take(p.id, start, end)) continue;
      const mm = mapLabel(label, SIGN_WORDS.test(label) ? { kind: 'applicant' } : group, `“${label}”에 맞는 입력 항목이 없습니다`);
      add({ kind: 'inline', paragraphId: p.id, range: { start, end }, label, group: group && group.text, where: whereP(p), field: mm.field, confidence: mm.confidence, why: mm.why, reason: `${p.cell ? '칸' : '문단'} 안 빈자리 · “${label}:”` });
    }
  });

  /* ---------- 3. 선택 표시 ---------- */
  const choiceUnits = []; // { paragraphs:[p], question, where, group }
  const handled = new Set();
  model.tables.forEach((table) => {
    table.cells.forEach((cell) => {
      const ps = cell.paragraphIds.map(paragraphOf).filter((p) => /\[\s*\]/.test(p.text || ''));
      if (ps.length === 0) return;
      const left = cell.col > 0 ? cellAt(table, cell.row, cell.col - 1) : null;
      const leftLabel = left && left !== cell && isLabelText(left.text) ? oneLine(left.text) : '';
      ps.forEach((p) => {
        handled.add(p.id);
        const before = (p.text.split('[')[0] || '').trim();
        if (before && !/^\(/.test(before) && before.length <= 40) {
          choiceUnits.push({ paragraphs: [p], question: oneLine(before.replace(/[:：]$/, '')), where: whereP(p), group: groupOf(table, cell) });
        } else {
          const last = choiceUnits[choiceUnits.length - 1];
          if (last && last.cellRef === cell) last.paragraphs.push(p);
          else choiceUnits.push({ paragraphs: [p], question: leftLabel || '(질문 없음)', where: whereP(p), group: groupOf(table, cell), cellRef: cell });
        }
      });
    });
  });
  model.paragraphs.forEach((p) => {
    if (handled.has(p.id) || !/\[\s*\]/.test(p.text || '')) return;
    const before = (p.text.split('[')[0] || '').trim();
    choiceUnits.push({ paragraphs: [p], question: oneLine(before.replace(/[:：]$/, '')) || '(질문 없음)', where: whereP(p), group: null });
  });

  choiceUnits.forEach((unit) => {
    const main = [];
    const subs = new Map(); // parenStart key -> { question, options, parentIndex }
    unit.paragraphs.forEach((p) => {
      const text = p.text;
      // 괄호 깊이를 따라가며 선택지를 모은다
      let depth = 0;
      let parenKey = null;
      let parenStart = -1;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '(' || ch === '（') {
          depth += 1;
          if (depth === 1) {
            parenKey = `${p.id}:${i}`;
            parenStart = i;
          }
          continue;
        }
        if (ch === ')' || ch === '）') {
          depth = Math.max(0, depth - 1);
          continue;
        }
        if (ch !== '[') continue;
        const close = text.indexOf(']', i);
        if (close < 0 || text.slice(i + 1, close).trim() !== '') continue;
        const rest = text.slice(close + 1);
        // 선택지 이름: 다음 빈 상자·괄호·쉼표·빗금·두 칸 이상 공백 앞까지("[소아]"처럼 글자가 든 대괄호는 이름에 포함)
        const optText = oneLine((/^\s*((?:\[[^\]\s][^\]]*\]|[^[\]()（）,/:：])*?)(?=\s{2,}|\s*\[\s*\]|\s*[()（）,/]|\s*$)/.exec(rest) || ['', ''])[1]);
        if (/^\s*(안에|에는|에\s|표시)/.test(rest)) {
          i = close; // 안내 문구("[  ]에는 해당되는 곳에 √ 표시") — 선택지가 아니다
          continue;
        }
        const option = { text: optText || '(이름 없음)', paragraphId: p.id, range: { start: i + 1, end: close } };
        if (depth === 0) {
          main.push(option);
        } else {
          if (!subs.has(parenKey)) {
            const head = text.slice(parenStart + 1, i).replace(/[:：]/g, '').trim();
            subs.set(parenKey, { question: head, options: [], parentIndex: main.length - 1 });
          }
          subs.get(parenKey).options.push(option);
        }
        i = close;
      }
    });
    if (main.length === 0 && subs.size === 0) return;
    const many = main.length >= 6;
    const parent = main.length
      ? add({ kind: 'choice', paragraphId: main[0].paragraphId, range: main[0].range, question: unit.question, label: unit.question, group: unit.group && unit.group.text, where: unit.where, options: main, multiple: many, field: null, confidence: 'sure', why: '', reason: `선택 표시 ${main.length}개${many ? ' · 여러 개 고를 수 있게 함' : ''}` })
      : null;
    subs.forEach((sub) => {
      const parentOption = parent && sub.parentIndex >= 0 ? parent.options[sub.parentIndex] : null;
      const question = sub.question || (parentOption ? `${parentOption.text} 세부` : unit.question);
      add({
        kind: 'choice',
        paragraphId: sub.options[0].paragraphId,
        range: sub.options[0].range,
        question,
        label: question,
        group: unit.group && unit.group.text,
        where: unit.where,
        options: sub.options,
        multiple: false,
        dependsOn: parent && sub.parentIndex >= 0 ? { spotId: parent.id, option: sub.parentIndex } : null,
        field: null,
        confidence: 'sure',
        why: '',
        reason: `괄호 안 하위 선택 ${sub.options.length}개${parentOption ? ` · “${parentOption.text}”을 고를 때` : ''}`,
      });
    });
  });

  /* ---------- 4. 선택지 뒤 괄호 안 빈자리는 그 선택지를 골랐을 때만 ---------- */
  const choices = spots.filter((s) => s.kind === 'choice');
  spots
    .filter((s) => s.kind === 'inline')
    .forEach((s) => {
      const p = paragraphOf(s.paragraphId);
      choices.forEach((c) => {
        c.options.forEach((o, idx) => {
          if (o.paragraphId !== s.paragraphId || o.range.end > s.range.start) return;
          const between = p.text.slice(o.range.end, s.range.start);
          if (/[(（]/.test(between) && !/[)）]/.test(between) && !/\[\s*\]/.test(between.slice(between.indexOf('(')))) {
            s.dependsOn = { spotId: c.id, option: idx };
            s.reason += ` · “${o.text}”을 고를 때`;
            // 맞는 입력 항목이 없는 조건부 빈자리(기관·금액 등)는 출력할 때 그 서식에서만 묻는다
            if (!s.field) {
              s.ask = true;
              s.confidence = 'sure';
              s.why = '';
            }
          }
        });
      });
    });

  // 문서 순서대로(문단 순서 → 글자 위치) 번호
  const order = new Map(model.paragraphs.map((p, i) => [p.id, i]));
  spots.sort((a, b) => order.get(a.paragraphId) - order.get(b.paragraphId) || a.range.start - b.range.start);
  spots.forEach((s, i) => {
    s.no = i + 1;
  });
  return spots;
}

/** 인식 결과 요약: { total, byKind, sure, check, none } */
function summarize(spots) {
  const byKind = {};
  spots.forEach((s) => (byKind[s.kind] = (byKind[s.kind] || 0) + 1));
  return {
    total: spots.length,
    byKind,
    sure: spots.filter((s) => s.confidence === 'sure' && (s.field || s.kind === 'choice')).length,
    check: spots.filter((s) => s.confidence === 'check').length,
    none: spots.filter((s) => s.confidence === 'none').length,
  };
}

module.exports = { recognize, summarize, isLabelText };
