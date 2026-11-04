// ui-dom.js
// 화면 요소를 만드는 작은 도우미. 글자는 늘 textContent로 넣어(HTML로 해석하지 않음) 서식 이름·입력값이 화면을 깨뜨리지 않게 한다.

/**
 * h('button', { class: 'fa-btn', onclick: fn, 'aria-pressed': 'true' }, '글자', 자식요소 …)
 * 속성: on* 함수는 이벤트, class·for·hidden·checked·value·disabled는 그 속성, 나머지는 setAttribute
 */
function h(tag, attrs = {}, ...children) {
  const el = document.createElement(tag);
  Object.entries(attrs || {}).forEach(([k, v]) => {
    if (v === undefined || v === null || v === false) return;
    if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else if (k === 'class') el.className = v;
    else if (k === 'for') el.htmlFor = v;
    else if (['hidden', 'checked', 'disabled', 'value', 'readOnly'].includes(k)) el[k] = v;
    else el.setAttribute(k, v === true ? '' : String(v));
  });
  append(el, children);
  return el;
}

function append(el, children) {
  children.flat(Infinity).forEach((c) => {
    if (c === null || c === undefined || c === false) return;
    el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  });
  return el;
}

/** 요소의 내용을 바꾼다 */
function replace(el, ...children) {
  el.replaceChildren();
  return append(el, children);
}

/** <select>: options는 [값] 또는 [[값, 보이는 글자]] · value는 고른 값 · onchange(값) */
function select(options, value, onchange, attrs = {}) {
  return h(
    'select',
    { ...attrs, onchange: (ev) => onchange(ev.target.value) },
    options.map((o) => {
      const [v, label] = Array.isArray(o) ? o : [o, o];
      return h('option', { value: v, selected: v === value ? 'selected' : null }, label);
    })
  );
}

module.exports = { h, append, replace, select };
