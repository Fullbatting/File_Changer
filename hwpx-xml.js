// hwpx-xml.js
// HWPX 본문 XML을 "원문 위치를 그대로 기억하는" 가벼운 트리로 읽고, 필요한 글자만 바꾼다.
// 일반 XML 라이브러리는 다시 쓸 때 속성 순서·공백·이름공간 표기를 바꿔 버릴 수 있다. 한글 서식은
// 그런 차이에도 민감할 수 있으므로, 원문 문자열에서 바꿀 구간만 잘라 끼우고 나머지는 한 글자도
// 건드리지 않는다. 한글이 저장한 OWPML(기계가 만든 정형 XML)을 대상으로 한다.

/** XML 문자 참조를 푼다 */
function decodeEntities(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    if (e[1] === 'x' || e[1] === 'X') return String.fromCodePoint(parseInt(e.slice(2), 16));
    return String.fromCodePoint(parseInt(e.slice(1), 10));
  });
}

/** 글자 노드에 넣을 수 있게 특수문자를 바꾼다 */
function encodeText(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const localName = (name) => name.slice(name.indexOf(':') + 1);

/**
 * xml을 트리로 읽는다. 요소: { type:'el', name, local, attrs(원문), start, openEnd, closeStart, end, children, parent }
 * 글자: { type:'text', start, end, parent }. start/end는 원문 xml의 위치(끝은 포함하지 않음).
 * 주석·처리 명령·CDATA·DOCTYPE은 건너뛴다(글자로 취급하지 않음).
 */
function parseXml(xml) {
  const root = { type: 'root', children: [], start: 0, end: xml.length };
  let cur = root;
  const re = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<[?!][\s\S]*?>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:"[^"]*"|'[^']*'|[^'">])*?)(\/?)>|[^<]+/g;
  let m;
  while ((m = re.exec(xml))) {
    const tok = m[0];
    if (tok[0] !== '<') {
      cur.children.push({ type: 'text', start: m.index, end: m.index + tok.length, parent: cur });
      continue;
    }
    if (tok.startsWith('<!--') || tok.startsWith('<?') || tok.startsWith('<!')) continue;
    if (m[1]) {
      // 닫는 태그: 같은 이름의 열린 요소까지 올라간다
      let el = cur;
      while (el && el.type === 'el' && el.name !== m[1]) el = el.parent;
      if (!el || el.type !== 'el') throw new Error(`XML 구조가 맞지 않습니다(닫는 태그 ${m[1]})`);
      el.closeStart = m.index;
      el.end = m.index + tok.length;
      cur = el.parent;
      continue;
    }
    const el = {
      type: 'el',
      name: m[2],
      local: localName(m[2]),
      attrs: m[3] || '',
      start: m.index,
      openEnd: m.index + tok.length,
      closeStart: null,
      end: null,
      children: [],
      parent: cur,
      selfClosing: m[4] === '/',
    };
    cur.children.push(el);
    if (el.selfClosing) {
      el.closeStart = el.openEnd;
      el.end = el.openEnd;
    } else {
      cur = el;
    }
  }
  if (cur !== root) throw new Error(`XML이 끝나지 않았습니다(${cur.name})`);
  return root;
}

/** 속성 값(풀어 쓴 글자). 없으면 null */
function attr(el, name) {
  const m = new RegExp(`(?:^|\\s)${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(el.attrs);
  if (!m) return null;
  return decodeEntities(m[1] !== undefined ? m[1] : m[2]);
}

/** 자식 중 로컬 이름이 local인 요소들 */
const childrenByLocal = (el, local) => el.children.filter((c) => c.type === 'el' && c.local === local);

/** 깊이 우선으로 모든 요소를 돈다. visit가 false를 돌려주면 그 아래로는 내려가지 않는다 */
function walk(el, visit) {
  el.children.forEach((c) => {
    if (c.type !== 'el') return;
    if (visit(c) === false) return;
    walk(c, visit);
  });
}

/** 가장 가까운 조상 중 로컬 이름이 local인 요소(없으면 null) */
function closest(el, local) {
  let p = el.parent;
  while (p && p.type === 'el') {
    if (p.local === local) return p;
    p = p.parent;
  }
  return null;
}

/**
 * 원문 xml에 여러 편집({ start, end, text })을 한꺼번에 적용한다. 편집 구간은 겹치면 안 된다.
 * 뒤에서부터 적용해 앞쪽 위치가 밀리지 않게 한다.
 */
function applyEdits(xml, edits) {
  const sorted = edits.slice().sort((a, b) => b.start - a.start || b.end - a.end);
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].end > sorted[i - 1].start) throw new Error('겹치는 편집이 있습니다');
  }
  let out = xml;
  sorted.forEach((e) => {
    out = out.slice(0, e.start) + e.text + out.slice(e.end);
  });
  return out;
}

module.exports = { parseXml, attr, childrenByLocal, walk, closest, applyEdits, decodeEntities, encodeText, localName };
