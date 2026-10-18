// mapping-workbench.js
// 탭1 ②단계 "매핑 작업대" 화면: 요약 바 · 연결 보기(목록 보기) · 결과 미리보기.
// 상태는 갖지 않는다 — tab1.js가 넘겨준 api로 상태를 읽고, 바꿀 때는 api의 조작 함수
// (내부적으로 commit 하나를 거침)를 부른다. 데이터 값은 항상 textContent로 넣는다.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const common = require('./common.js');

const {
  formatPreviewValue,
  matchesColumnSearch,
  rankSourceCandidates,
  showConfirmDialog,
  getUserDataPath,
  onUserDataPathReady,
  safeReadFileSync,
  safeWriteFileSync,
  log,
} = common;

const UI_SETTINGS_FILE_NAME = 'ui-settings.json';

const REASON_LABELS = {
  exact: '같은 이름',
  synonym: '동의어',
  fuzzy: '비슷한 이름',
  confirmed: '확인됨',
  manual: '직접 연결',
  preset: '프리셋',
  rule: '매핑 규칙',
  none: '연결 없음',
};

const ROW_H = 60; // 연결 보기 한 줄(카드 52 + 간격 8)
const HEAD_H = 40; // 묶음 머리
const COLLAPSED_H = 56; // 접힌 묶음 머리
const VIRTUAL_OVERSCAN = 360; // 보이는 범위 위아래로 미리 그려 둘 높이(px)
const GROUP_THRESHOLD = 12; // 타겟 컬럼이 이보다 많으면 기본으로 상태별 묶음 · 열 지도 표시
const PIN_THRESHOLD = 8; // 타겟 컬럼이 이보다 많으면 첫 열 고정 · 열 거르기 표시
const PREVIEW_PAGE = 50;
const COL_W = 150;
const ROWNUM_W = 92;
const H_VIRTUAL_THRESHOLD = 30; // 가로로 스크롤되는 열이 이보다 많으면 보이는 열만 그린다
const LIST_LAZY_OPTIONS = 20000;
const LIST_ALWAYS_RENDER = 60;
const LIST_BATCH = 60; // 목록 보기를 한 번에 그리는 줄 수 // 타겟 컬럼이 이 이하면 목록 보기가 숨겨져 있어도 함께 그린다 // 목록 보기 드롭다운 항목 수(행×열)가 이보다 많으면 열 때 채운다
const RECOMMEND_MIN_SCORE = 30;
const TOAST_MS = 5000;

const ICONS = {
  undo: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/></svg>',
  redo: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/></svg>',
  search: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  chevDown: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m6 9 6 6 6-6"/></svg>',
  chevRight: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9 6 6 6-6 6"/></svg>',
  x: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  grip: '<svg viewBox="0 0 12 16" aria-hidden="true" class="fill"><circle cx="3" cy="3" r="1.5"/><circle cx="9" cy="3" r="1.5"/><circle cx="3" cy="8" r="1.5"/><circle cx="9" cy="8" r="1.5"/><circle cx="3" cy="13" r="1.5"/><circle cx="9" cy="13" r="1.5"/></svg>',
  pin: '<svg viewBox="0 0 24 24" aria-hidden="true" class="fill"><path d="M15 3l6 6-3 1-4 4 1 5-2 2-4-4-5 5-1-1 5-5-4-4 2-2 5 1 4-4z"/></svg>',
  info: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.01"/></svg>',
};

/** 작은 DOM 생성 도우미. text는 textContent로, html은 정해진 아이콘 문자열에만 쓴다 */
function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  Object.entries(props).forEach(([key, value]) => {
    if (value === undefined || value === null || value === false) return;
    if (key === 'className') node.className = value;
    else if (key === 'text') node.textContent = value;
    else if (key === 'html') node.innerHTML = value;
    else if (key === 'dataset') Object.entries(value).forEach(([k, v]) => (node.dataset[k] = v));
    else if (key === 'style') node.style.cssText = value;
    else if (key.startsWith('on')) node.addEventListener(key.slice(2), value);
    else node.setAttribute(key, value === true ? '' : value);
  });
  children.flat().forEach((child) => {
    if (child === null || child === undefined || child === false) return;
    node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
  });
  return node;
}

const icon = (name, className = 'wb-icon') => el('span', { className, html: ICONS[name] });
const isProblemReason = (reason, mapped) => reason === 'fuzzy' || !mapped;

function createWorkbench(api) {
  const $ = (id) => document.getElementById(id);
  const root = {
    empty: $('wbEmpty'),
    main: $('wbMain'),
    banner: $('wbBanner'),
    summary: $('wbSummary'),
    connect: $('wbConnect'),
    listView: $('wbListView'),
    listBody: $('mappingTableBody'),
    preview: $('wbPreview'),
    offNote: $('wbOffNote'),
    visualToggle: $('wbVisualToggle'),
    visualState: $('wbVisualState'),
  };

  const ui = {
    visual: false, // 시각화(연결 보기·결과 미리보기·값 고치기). 기본은 꺼짐 — 끄면 매핑 표(목록 보기)만
    view: 'list',
    groupMode: 'order',
    collapsed: { auto: true },
    search: '',
    traySearch: '',
    showEmptyTray: false,
    focusedB: null,
    items: [],
    offsets: [],
    totalHeight: 0,
    picker: null,
    recCache: new Map(),
    renderCallbacks: [],
    // 화면 크기·스크롤 위치는 ResizeObserver/scroll 이벤트로만 갱신해 둔다. 그리는 도중에
    // clientHeight·scrollLeft를 읽으면 방금 바꾼 DOM 때문에 강제 레이아웃이 일어나 느려진다.
    layout: { rowsTop: 0, rowsHeight: 560, pvLeft: 0, pvWidth: 1000 },
    preview: {
      page: 0,
      rowFilter: false,
      rowFilterList: null,
      colFilter: false,
      pinned: new Set(),
      active: null, // { i, b }
      editing: null, // { i, b }
      thByB: new Map(),
      range: null,
    },
  };

  let st = null; // 그릴 때마다 api.getState()로 새로 읽는다

  function setRowsScrollTop(value) {
    const max = Math.max(0, ui.totalHeight - ui.layout.rowsHeight);
    ui.layout.rowsTop = Math.max(0, Math.min(max, value));
    rowsScroller.scrollTop = ui.layout.rowsTop;
  }

  function setPreviewScrollLeft(value) {
    ui.layout.pvLeft = Math.max(0, value);
    pvWrap.scrollLeft = ui.layout.pvLeft;
  }

  /* =====================================================================
   * 연결 보기 뼈대 (한 번만 만든다)
   * ===================================================================== */

  const searchInput = el('input', {
    type: 'text',
    id: 'wbSearch',
    placeholder: '컬럼 찾기 · 이름, 초성(ㅇㄹㅊ), 값',
    'aria-label': '컬럼 찾기',
  });
  const groupStatusBtn = el('button', { type: 'button', id: 'wbGroupStatus', text: '상태별' });
  const groupOrderBtn = el('button', { type: 'button', id: 'wbGroupOrder', text: '양식 순서' });
  const connHeadA = el('span', { className: 'wb-col-head' });
  const connHeadB = el('span', { className: 'wb-col-head' });
  const rowsLayer = el('div', { className: 'wb-rows-layer' });
  const rowsScroller = el('div', { className: 'wb-rows', id: 'wbRows', role: 'list', 'aria-label': '컬럼 연결 목록' }, rowsLayer);
  const trayCount = el('span', { className: 'wb-count' });
  const traySearch = el('input', { type: 'text', id: 'wbTraySearch', placeholder: '이 목록에서 찾기', 'aria-label': '연결 안 된 원본 컬럼 찾기' });
  const trayList = el('div', { className: 'wb-tray-list', id: 'wbTrayList' });
  const trayFooter = el('div', { className: 'wb-tray-footer' });
  const tray = el(
    'aside',
    { className: 'wb-tray', id: 'wbTray', 'aria-label': '연결 안 된 원본 컬럼' },
    el('div', { className: 'wb-tray-head' }, el('b', { text: '연결 안 된 원본 컬럼' }), trayCount),
    el('label', { className: 'wb-search small' }, icon('search'), traySearch),
    el('div', { className: 'wb-hint', text: '타겟 쪽 빈 자리로 끌어다 놓으면 연결됩니다' }),
    trayList,
    trayFooter
  );

  root.connect.append(
    el(
      'div',
      { className: 'wb-toolbar' },
      el('label', { className: 'wb-search' }, icon('search'), searchInput),
      el('span', { className: 'wb-muted', text: '묶어 보기' }),
      el('div', { className: 'wb-seg', role: 'group', 'aria-label': '묶어 보기' }, groupStatusBtn, groupOrderBtn),
      el('span', { className: 'wb-spacer' }),
      el(
        'span',
        { className: 'wb-keys' },
        el('kbd', { text: 'J' }),
        el('kbd', { text: 'K' }),
        ' 다음·이전 확인 항목 ',
        el('kbd', { text: 'Y' }),
        ' 맞음 ',
        el('kbd', { text: 'C' }),
        ' 바꾸기'
      )
    ),
    el(
      'div',
      { className: 'wb-connect-body' },
      el('div', { className: 'wb-rows-col' }, el('div', { className: 'wb-col-heads' }, connHeadA, connHeadB), rowsScroller),
      tray
    )
  );

  searchInput.addEventListener('input', () => {
    ui.search = searchInput.value;
    setRowsScrollTop(0);
    renderConnect();
  });
  traySearch.addEventListener('input', () => {
    ui.traySearch = traySearch.value;
    renderTray();
  });
  groupStatusBtn.addEventListener('click', () => setGroupMode('status'));
  groupOrderBtn.addEventListener('click', () => setGroupMode('order'));

  function setGroupMode(mode) {
    ui.groupMode = mode;
    setRowsScrollTop(0);
    renderConnect();
  }

  /* =====================================================================
   * 결과 미리보기 뼈대
   * ===================================================================== */

  const pvRange = el('span', { className: 'wb-muted wb-num', id: 'wbPreviewRange' });
  const pvColSearch = el('input', { type: 'text', id: 'wbColJump', placeholder: '열로 이동', 'aria-label': '열로 이동' });
  const pvColSearchWrap = el('label', { className: 'wb-search small' }, icon('search'), pvColSearch);
  const pvColAll = el('button', { type: 'button', id: 'wbColAll', text: '모든 열' });
  const pvColProblem = el('button', { type: 'button', id: 'wbColProblem' });
  const pvColSeg = el('div', { className: 'wb-seg', role: 'group', 'aria-label': '열 거르기' }, pvColAll, pvColProblem);
  const pvRowFilter = el('button', { type: 'button', className: 'wb-btn', id: 'wbRowFilter', text: '문제 있는 행만' });
  const pvColMap = el('div', { className: 'wb-colmap', id: 'wbColMap' });
  const pvHead = el('thead');
  const pvBody = el('tbody');
  const pvTable = el('table', { className: 'wb-grid', id: 'wbGrid' }, pvHead, pvBody);
  const pvWrap = el('div', { className: 'wb-grid-wrap', id: 'wbGridWrap' }, pvTable);
  const pvPrev = el('button', { type: 'button', className: 'wb-btn', id: 'wbPrevPage', text: '‹ 이전 50행' });
  const pvNext = el('button', { type: 'button', className: 'wb-btn', id: 'wbNextPage', text: '다음 50행 ›' });
  const pvFooterNote = el('span', { className: 'wb-muted' });
  const pvEmpty = el('div', { className: 'wb-pv-empty', hidden: true });

  root.preview.append(
    el(
      'div',
      { className: 'wb-pv-toolbar' },
      el('h3', { text: '결과 미리보기' }),
      el('span', { className: 'wb-muted', text: '저장될 표 그대로 · 칸을 더블클릭하거나 Enter로 고칠 수 있습니다' }),
      el('span', { className: 'wb-spacer' }),
      pvRange,
      pvColSearchWrap,
      pvColSeg,
      pvRowFilter
    ),
    pvColMap,
    pvWrap,
    pvEmpty,
    el('div', { className: 'wb-pv-footer' }, pvFooterNote, el('span', { className: 'wb-spacer' }), pvPrev, pvNext)
  );

  pvColAll.addEventListener('click', () => {
    ui.preview.colFilter = false;
    setPreviewScrollLeft(0);
    renderPreview();
  });
  pvColProblem.addEventListener('click', () => {
    ui.preview.colFilter = true;
    setPreviewScrollLeft(0);
    renderPreview();
  });
  pvRowFilter.addEventListener('click', () => setRowFilter(!ui.preview.rowFilter));
  pvPrev.addEventListener('click', () => {
    ui.preview.page = Math.max(0, ui.preview.page - 1);
    renderPreview();
  });
  pvNext.addEventListener('click', () => {
    ui.preview.page += 1;
    renderPreview();
  });
  pvColSearch.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const q = pvColSearch.value.trim();
    if (!q || !st) return;
    const b = st.headersB.find((h) => matchesColumnSearch(q, api.labelOfB(h), []));
    if (b) scrollPreviewToColumn(b, { flash: true });
  });

  let pvScrollFrame = 0;
  pvWrap.addEventListener('scroll', () => {
    ui.layout.pvLeft = pvWrap.scrollLeft;
    hideBubble();
    if (pvScrollFrame) return;
    pvScrollFrame = requestAnimationFrame(() => {
      pvScrollFrame = 0;
      if (!st || ui.preview.editing) return;
      const before = ui.preview.range;
      const next = computeColumnRange();
      if (!before || before.first !== next.first || before.last !== next.last) renderPreview();
      else updateColMapViewport();
    });
  });

  /* ---------- 토스트 · 고친 값 말풍선 ---------- */

  const toastText = el('span');
  const toastUndo = el('button', { type: 'button', text: '실행 취소' });
  const toast = el('div', { className: 'wb-toast', id: 'wbToast', role: 'status', hidden: true }, toastText, toastUndo);
  document.body.appendChild(toast);
  let toastTimer = 0;
  toastUndo.addEventListener('click', () => {
    toast.hidden = true;
    api.undo();
  });
  function showToast(text) {
    toastText.textContent = text;
    toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (toast.hidden = true), TOAST_MS);
  }

  const bubbleText = el('div', { className: 'wb-bubble-text' });
  const bubbleNow = el('div', { className: 'wb-bubble-text' });
  const bubbleRevert = el('button', { type: 'button', className: 'wb-btn', id: 'wbRevertCell', text: '원래 값으로' });
  const bubble = el(
    'div',
    { className: 'wb-bubble', id: 'wbBubble', hidden: true },
    el('div', { className: 'wb-bubble-title', text: '직접 고친 값' }),
    bubbleText,
    bubbleNow,
    bubbleRevert
  );
  document.body.appendChild(bubble);
  bubbleRevert.addEventListener('mousedown', (e) => e.preventDefault());
  bubbleRevert.addEventListener('click', async () => {
    const target = bubble.target;
    hideBubble();
    if (!target) return;
    const ok = await api.clearOverride(target.i, target.b);
    if (ok) focusCell(target.i, target.b);
  });
  function hideBubble() {
    bubble.hidden = true;
    bubble.target = null;
  }
  function showBubble(td, i, b) {
    const text = api.getOverride(i, b);
    if (text === undefined) return hideBubble();
    const original = api.displayValue(api.originalValue(i, b), b);
    bubbleText.textContent = `원래 값: ${original === '' ? '(빈칸)' : original}`;
    bubbleNow.textContent = `지금 값: ${text === '' ? '(빈칸)' : text}`;
    const rect = td.getBoundingClientRect();
    bubble.hidden = false;
    bubble.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - bubble.offsetWidth - 8))}px`;
    const below = rect.bottom + 6;
    bubble.style.top = `${below + bubble.offsetHeight > window.innerHeight - 8 ? rect.top - bubble.offsetHeight - 6 : below}px`;
    bubble.target = { i, b };
  }

  /* =====================================================================
   * 전체 그리기
   * ===================================================================== */

  function render() {
    st = api.getState();
    hideBubble();
    if (!st.source) {
      root.empty.hidden = false;
      root.main.hidden = true;
      closePicker();
      renderList();
      ui.renderCallbacks.forEach((cb) => cb());
      return;
    }
    root.empty.hidden = true;
    root.main.hidden = false;
    renderBanner();
    renderOffNote();
    if (!ui.visual) {
      // 시각화 꺼짐: 매핑 표만
      root.summary.hidden = true;
      root.connect.hidden = true;
      root.preview.hidden = true;
      root.listView.hidden = false;
      renderList();
      ui.renderCallbacks.forEach((cb) => cb());
      return;
    }
    root.summary.hidden = false;
    root.preview.hidden = false;
    renderSummary();
    root.connect.hidden = ui.view !== 'connect';
    root.listView.hidden = ui.view !== 'list';
    if (ui.view === 'connect') renderConnect();
    // 목록 보기는 보일 때만 다시 그린다. 단, 컬럼이 적으면(가벼우면) 숨겨져 있어도 늘 최신으로 둔다
    if (ui.view === 'list' || st.headersB.length <= LIST_ALWAYS_RENDER) renderList();
    renderPreview();
    if (ui.picker) renderPickerList();
    ui.renderCallbacks.forEach((cb) => cb());
  }

  /* ---------- 시각화 켜기/끄기 ---------- */

  function settingsPath() {
    return path.join(getUserDataPath(), UI_SETTINGS_FILE_NAME);
  }

  function syncVisualToggle() {
    root.visualToggle.checked = ui.visual;
    root.visualState.textContent = ui.visual ? 'ON' : 'OFF';
  }

  function setVisual(on, { save = true } = {}) {
    ui.visual = !!on;
    ui.view = ui.visual ? 'connect' : 'list';
    closePicker();
    hideBubble();
    toast.hidden = true;
    syncVisualToggle();
    if (save) {
      try {
        safeWriteFileSync(settingsPath(), JSON.stringify({ visual: ui.visual }, null, 2));
      } catch (err) {
        log(`화면 설정을 저장하지 못했습니다: ${err.message}`, 'warn');
      }
    }
    render();
  }

  root.visualToggle.addEventListener('change', () => setVisual(root.visualToggle.checked));

  // 사용자 데이터 폴더가 준비되면 지난번 켜기/끄기 상태를 불러온다(파일이 없으면 기본값 꺼짐)
  onUserDataPathReady(() => {
    let visual = false;
    try {
      if (fs.existsSync(settingsPath())) visual = JSON.parse(safeReadFileSync(settingsPath(), 'utf-8')).visual === true;
    } catch (err) {
      log(`화면 설정을 불러오지 못해 기본값(시각화 꺼짐)으로 시작합니다: ${err.message}`, 'warn');
    }
    setVisual(visual, { save: false });
  });
  syncVisualToggle();

  /** 시각화를 끈 상태에서도 직접 고친 값이 저장에 들어간다는 것을 알린다 */
  function renderOffNote() {
    const edited = st.overrides.size;
    root.offNote.hidden = ui.visual || edited === 0;
    if (root.offNote.hidden) return;
    root.offNote.textContent = `직접 고친 값 ${edited}개가 저장에 포함됩니다. 확인하거나 되돌리려면 시각화를 켜세요.`;
  }

  /* ---------- 프리셋 제안 띠 ---------- */

  function renderBanner() {
    const suggestion = api.getPresetSuggestion();
    root.banner.innerHTML = '';
    root.banner.hidden = !suggestion;
    if (!suggestion) return;
    root.banner.append(
      icon('info'),
      el(
        'span',
        {},
        '저장된 프리셋 ',
        el('b', { text: suggestion.preset.name }),
        '이(가) 이 양식과 ',
        el('b', { text: `${suggestion.total}열 중 ${suggestion.matched}열` }),
        ' 같습니다.'
      ),
      el('span', { className: 'wb-spacer' }),
      el('button', {
        type: 'button',
        className: 'btn primary small',
        id: 'wbApplyPreset',
        text: '프리셋 적용',
        onclick: () => api.applyPreset(suggestion.preset.id),
      }),
      el('button', { type: 'button', className: 'btn ghost small', id: 'wbDismissPreset', text: '닫기', onclick: () => api.dismissPresetSuggestion() })
    );
  }

  /* ---------- 요약 바 ---------- */

  function summaryCounts() {
    const derived = api.getDerived();
    const total = st.headersB.length;
    const connected = st.headersB.filter((b) => st.mapping[b]).length;
    const fuzzy = st.headersB.filter((b) => st.matchReason[b] === 'fuzzy').length;
    const none = total - connected;
    const requiredNone = st.headersB.filter((b) => !st.mapping[b] && st.required.has(b)).length;
    return {
      total,
      connected,
      fuzzy,
      none,
      requiredNone,
      missingRows: derived.missingRowCount,
      duplicateRows: derived.duplicateCount,
      edited: st.overrides.size,
    };
  }

  function chip(id, text, tone, onClick) {
    return el('button', { type: 'button', id, className: `wb-chip-stat ${tone}`, text, onclick: onClick });
  }

  function renderSummary() {
    const c = summaryCounts();
    root.summary.innerHTML = '';
    const parts = [
      chip('wbStatConnected', `연결 ${c.connected} / ${c.total}`, c.connected === c.total ? 'ok' : 'neutral', () => setView('connect')),
      chip('wbStatFuzzy', `확인 필요 ${c.fuzzy}`, c.fuzzy ? 'warn' : 'neutral', () => jumpToProblem('fuzzy')),
      chip('wbStatNone', `연결 없음 ${c.none}${c.requiredNone ? ` · 필수 ${c.requiredNone}` : ''}`, c.requiredNone ? 'error' : c.none ? 'warn-soft' : 'neutral', () =>
        jumpToProblem('none')
      ),
      chip('wbStatMissing', `필수 빈칸 ${c.missingRows.toLocaleString()}행`, c.missingRows ? 'error' : 'neutral', () => showProblemRows()),
      chip('wbStatDuplicate', `중복 ${c.duplicateRows.toLocaleString()}행`, c.duplicateRows ? 'error' : 'neutral', () => showProblemRows()),
      chip('wbStatEdited', `고친 값 ${c.edited}개`, c.edited ? 'edited' : 'neutral', () => showProblemRows()),
      el('span', { className: 'wb-spacer' }),
      el('button', {
        type: 'button',
        className: 'wb-icon-btn',
        id: 'wbUndo',
        title: '실행 취소 (Ctrl+Z)',
        'aria-label': '실행 취소',
        disabled: !st.canUndo,
        html: ICONS.undo,
        onclick: () => api.undo(),
      }),
      el('button', {
        type: 'button',
        className: 'wb-icon-btn',
        id: 'wbRedo',
        title: '다시 실행 (Ctrl+Y)',
        'aria-label': '다시 실행',
        disabled: !st.canRedo,
        html: ICONS.redo,
        onclick: () => api.redo(),
      }),
      c.edited
        ? el('button', { type: 'button', className: 'wb-btn', id: 'wbRevertAll', text: '모두 되돌리기', onclick: () => api.clearAllOverrides() })
        : null,
      el(
        'div',
        { className: 'wb-seg', role: 'group', 'aria-label': '보기 방식' },
        el('button', {
          type: 'button',
          id: 'wbViewConnect',
          text: '연결 보기',
          'aria-pressed': String(ui.view === 'connect'),
          onclick: () => setView('connect'),
        }),
        el('button', {
          type: 'button',
          id: 'wbViewList',
          text: '목록 보기',
          'aria-pressed': String(ui.view === 'list'),
          onclick: () => setView('list'),
        })
      ),
    ];
    root.summary.append(...parts.filter(Boolean));
  }

  function setView(view) {
    if (!ui.visual || ui.view === view) return;
    ui.view = view;
    closePicker();
    render();
  }

  function showProblemRows() {
    if (!ui.preview.rowFilter) setRowFilter(true);
    root.preview.scrollIntoView({ block: 'nearest' });
  }

  /* =====================================================================
   * 연결 보기: 한 줄 = 타겟 컬럼 하나, [원본 카드]—선—[타겟 카드]
   * ===================================================================== */

  function usageCount() {
    const counts = new Map();
    st.headersB.forEach((b) => {
      const a = st.mapping[b];
      if (a) counts.set(a, (counts.get(a) || 0) + 1);
    });
    return counts;
  }

  function rowMatchesSearch(b) {
    const q = ui.search.trim();
    if (!q) return true;
    if (matchesColumnSearch(q, api.labelOfB(b), [])) return true;
    const a = st.mapping[b];
    return !!a && matchesColumnSearch(q, a, st.source.samples[a]);
  }

  function groupOf(b) {
    const reason = st.matchReason[b] || 'none';
    if (!st.mapping[b]) return 'none';
    if (reason === 'fuzzy') return 'fuzzy';
    if (reason === 'exact' || reason === 'synonym') return 'auto';
    return 'manual';
  }

  /** 표시할 항목(묶음 머리·줄) 목록과 각 항목의 위치를 만든다 — 가상 스크롤의 기준 */
  function buildItems() {
    const visible = st.headersB.filter(rowMatchesSearch);
    const items = [];
    if (ui.groupMode === 'order') {
      visible.forEach((b) => items.push({ type: 'row', b }));
    } else {
      const groups = { fuzzy: [], none: [], manual: [], auto: [] };
      visible.forEach((b) => groups[groupOf(b)].push(b));
      groups.none.sort((x, y) => Number(st.required.has(y)) - Number(st.required.has(x)));
      const searching = ui.search.trim() !== '';
      ['fuzzy', 'none', 'manual', 'auto'].forEach((key) => {
        const list = groups[key];
        if (list.length === 0) return;
        const collapsed = !searching && !!ui.collapsed[key];
        items.push({ type: 'head', group: key, list, collapsed });
        if (!collapsed) list.forEach((b) => items.push({ type: 'row', b }));
      });
    }
    let y = 0;
    const offsets = items.map((item) => {
      const top = y;
      y += item.type === 'row' ? ROW_H : item.collapsed ? COLLAPSED_H : HEAD_H;
      return top;
    });
    ui.items = items;
    ui.offsets = offsets;
    ui.totalHeight = y;
  }

  function renderConnect() {
    if (!st || !st.source || !ui.visual) return;
    groupStatusBtn.setAttribute('aria-pressed', String(ui.groupMode === 'status'));
    groupOrderBtn.setAttribute('aria-pressed', String(ui.groupMode === 'order'));
    connHeadA.textContent = '';
    connHeadA.append(el('b', { text: '원본(A)' }), ` · ${st.sourceName} · ${st.headersA.length}열`);
    connHeadB.textContent = '';
    connHeadB.append(el('b', { text: '타겟(B)' }), ` · ${st.targetName} · ${st.headersB.length}열`);
    buildItems();
    rowsLayer.style.height = `${Math.max(ui.totalHeight, 1)}px`;
    renderConnectWindow();
    renderTray();
  }

  let connScrollFrame = 0;
  rowsScroller.addEventListener('scroll', () => {
    ui.layout.rowsTop = rowsScroller.scrollTop;
    if (connScrollFrame) return;
    connScrollFrame = requestAnimationFrame(() => {
      connScrollFrame = 0;
      renderConnectWindow();
    });
  });

  /** 보이는 범위(+여유)의 항목만 DOM으로 만든다 */
  function renderConnectWindow() {
    if (!st || !st.source || !ui.visual) return;
    const hadFocus = rowsScroller.contains(document.activeElement) && document.activeElement !== rowsScroller;
    const focusedKey = hadFocus ? document.activeElement.closest('[data-item-key]') : null;
    const restoreKey = focusedKey ? focusedKey.dataset.itemKey : null;

    const top = ui.layout.rowsTop;
    const height = ui.layout.rowsHeight;
    const from = top - VIRTUAL_OVERSCAN;
    const to = top + height + VIRTUAL_OVERSCAN;
    rowsLayer.innerHTML = '';
    const counts = usageCount();
    const used = new Set(counts.keys());
    let tabStopSet = false;
    const focusTarget = ui.focusedB && ui.items.some((it) => it.type === 'row' && it.b === ui.focusedB) ? ui.focusedB : null;

    if (ui.items.length === 0) {
      rowsLayer.appendChild(el('div', { className: 'wb-empty-row', text: ui.search.trim() ? '찾는 컬럼이 없습니다.' : '타겟 컬럼이 없습니다.' }));
      return;
    }

    ui.items.forEach((item, idx) => {
      const y = ui.offsets[idx];
      const h = item.type === 'row' ? ROW_H : item.collapsed ? COLLAPSED_H : HEAD_H;
      if (y + h < from || y > to) return;
      let node;
      if (item.type === 'head') {
        node = buildGroupHead(item);
      } else {
        node = buildRow(item.b, counts, used);
        const isStop = focusTarget ? item.b === focusTarget : !tabStopSet;
        if (isStop) {
          node.tabIndex = 0;
          tabStopSet = true;
        }
      }
      node.style.top = `${y}px`;
      rowsLayer.appendChild(node);
    });

    if (restoreKey) {
      const again = rowsLayer.querySelector(`[data-item-key="${CSS.escape(restoreKey)}"]`);
      if (again) again.focus({ preventScroll: true });
    }
  }

  const GROUP_TITLES = {
    fuzzy: ['확인 필요', '이름만 비슷해서 연결했습니다'],
    none: ['연결 없음', ''],
    manual: ['직접 연결', '직접 고른 연결 · 확인한 연결 · 프리셋/규칙'],
    auto: ['자동 연결됨', ''],
  };

  function buildGroupHead(item) {
    const [title, baseSub] = GROUP_TITLES[item.group];
    let sub = baseSub;
    if (item.group === 'none') {
      const req = item.list.filter((b) => st.required.has(b)).length;
      sub = req ? `필수 ${req}개` : '';
    } else if (item.group === 'auto') {
      const exact = item.list.filter((b) => st.matchReason[b] === 'exact').length;
      sub = `같은 이름 ${exact} · 동의어 ${item.list.length - exact}${item.collapsed ? ' · 값도 확인했으면 접어 두세요' : ''}`;
    }
    const toggle = el(
      'button',
      {
        type: 'button',
        className: 'wb-group-toggle',
        'aria-expanded': String(!item.collapsed),
        dataset: { itemKey: `head:${item.group}`, groupToggle: item.group },
      },
      icon(item.collapsed ? 'chevRight' : 'chevDown'),
      el('b', { text: title }),
      el('span', { className: 'wb-count', text: String(item.list.length) }),
      sub ? el('span', { className: 'wb-muted', text: sub }) : null
    );
    const head = el('div', { className: `wb-group-head${item.collapsed ? ' collapsed' : ''}`, dataset: { group: item.group } }, toggle, el('span', { className: 'wb-spacer' }));
    if (item.group === 'fuzzy') {
      head.appendChild(
        el('button', { type: 'button', className: 'wb-btn warn', dataset: { action: 'confirm-all' }, id: 'wbConfirmAll', text: `남은 ${item.list.length}개 모두 맞음` })
      );
    }
    if (item.collapsed) {
      head.appendChild(el('button', { type: 'button', className: 'wb-btn', dataset: { groupToggle: item.group }, text: '펼치기' }));
    }
    return head;
  }

  /** 원본 카드의 값 예시. 숫자는 원본의 숫자 표시 형식(천 단위·% 등)으로 보여준다 */
  function sampleText(a, count = 3) {
    const format = st.source.numberFormats[a];
    return (st.source.samples[a] || [])
      .slice(0, count)
      .map((v) => {
        if (typeof v === 'number' && format) {
          try {
            return XLSX.SSF.format(format, v);
          } catch (e) {
            return formatPreviewValue(v);
          }
        }
        return formatPreviewValue(v);
      })
      .join(' · ');
  }

  /** b에 추천할 원본 컬럼(다른 곳에 연결 안 된 것, 값이 있는 것, 점수 기준 이상) */
  function recommendFor(b, limit, used) {
    if (!ui.recCache.has(b)) {
      ui.recCache.set(b, rankSourceCandidates(api.labelOfB(b), st.headersA, st.source.valueKinds));
    }
    return ui.recCache
      .get(b)
      .filter((c) => c.score >= RECOMMEND_MIN_SCORE && !used.has(c.a) && c.a !== st.mapping[b] && !st.source.emptyColumns.has(c.a))
      .slice(0, limit);
  }

  function buildRow(b, counts, used) {
    const a = st.mapping[b];
    const reason = a ? st.matchReason[b] || 'manual' : 'none';
    const required = st.required.has(b);
    const label = api.labelOfB(b);

    let left;
    if (a) {
      const n = counts.get(a) || 0;
      left = el(
        'div',
        { className: 'wb-a-card', dataset: { dragA: a, pick: b }, title: '끌어서 다른 타겟 컬럼에도 연결할 수 있습니다' },
        el('span', { className: 'wb-card-text' }, el('b', { text: a }), n > 1 ? el('span', { className: 'wb-tag', text: `${n}곳` }) : null, el('span', { className: 'wb-sample', text: sampleText(a) }))
      );
    } else {
      const ambiguous = (st.autoCandidates[b] || []).filter((x) => x !== a);
      let content;
      if (ambiguous.length > 1) {
        content = [el('span', { className: 'wb-muted', text: `후보 ${ambiguous.length}개가 비슷함` }), ...ambiguous.slice(0, 3).map((x) => candidateChip(x, '자동 연결 후보'))];
      } else {
        const recs = recommendFor(b, 2, used);
        content = recs.length
          ? [el('span', { className: 'wb-muted', text: '추천' }), ...recs.map((r) => candidateChip(r.a, r.reasons.join(' · ')))]
          : [el('span', { className: 'wb-muted', text: '원본 컬럼 고르기 · 비워 두면 빈 열로 저장' })];
      }
      left = el('div', { className: `wb-a-empty${required ? ' required' : ''}`, dataset: { dropB: b, pick: b } }, ...content);
    }

    const line = el(
      'div',
      { className: `wb-line reason-${reason}` },
      el('span', { className: 'wb-dot' }),
      a ? el('span', { className: 'wb-wire', html: '<svg viewBox="0 0 100 8" preserveAspectRatio="none" aria-hidden="true"><path d="M0 4 L100 4"/></svg>' }) : el('span', { className: 'wb-wire empty' }),
      el('span', { className: 'wb-dot' })
    );

    const actions = [];
    if (reason === 'fuzzy') {
      actions.push(el('button', { type: 'button', className: 'wb-btn warn small', dataset: { action: 'confirm' }, text: '맞음' }));
      actions.push(el('button', { type: 'button', className: 'wb-btn small', dataset: { action: 'pick' }, text: '바꾸기' }));
    }
    if (a) {
      actions.push(el('button', { type: 'button', className: 'wb-x', dataset: { action: 'disconnect' }, title: '연결 끊기', 'aria-label': `${label} 연결 끊기`, html: ICONS.x }));
    }

    const right = el(
      'div',
      { className: `wb-b-card reason-${reason}${required && !a ? ' required' : ''}`, dataset: { dropB: b, pick: b } },
      el('b', { className: 'wb-b-label', text: label, title: label !== b ? `컬럼키: ${b}` : null }),
      el('button', {
        type: 'button',
        className: `wb-req${required ? ' on' : ''}`,
        dataset: { action: 'required' },
        'aria-pressed': String(required),
        title: '필수 컬럼: 비어 있는 행을 알려줍니다',
        text: '필수',
      }),
      el('span', { className: `wb-badge reason-${reason}`, text: REASON_LABELS[reason] }),
      el('span', { className: 'wb-spacer' }),
      ...actions
    );

    return el(
      'div',
      {
        className: `wb-row${b === ui.focusedB ? ' focused' : ''}`,
        role: 'listitem',
        tabindex: '-1',
        'aria-label': `${label}: ${a ? `${a} 연결, ${REASON_LABELS[reason]}` : '연결 없음'}${required ? ', 필수' : ''}`,
        dataset: { b, itemKey: `row:${b}` },
      },
      left,
      line,
      right
    );
  }

  function candidateChip(a, title) {
    return el('button', { type: 'button', className: 'wb-cand', dataset: { connectA: a }, title, text: a });
  }

  /* ---------- 연결 안 된 원본 컬럼 트레이 ---------- */

  function renderTray() {
    const used = new Set(usageCount().keys());
    const unused = st.headersA.filter((a) => !used.has(a));
    const nonEmpty = unused.filter((a) => !st.source.emptyColumns.has(a));
    const empty = unused.filter((a) => st.source.emptyColumns.has(a));
    const shown = (ui.showEmptyTray ? unused : nonEmpty).filter((a) => matchesColumnSearch(ui.traySearch, a, st.source.samples[a]));
    trayCount.textContent = String(nonEmpty.length + (ui.showEmptyTray ? empty.length : 0));
    trayList.innerHTML = '';
    if (shown.length === 0) {
      trayList.appendChild(el('div', { className: 'wb-muted wb-tray-none', text: unused.length === 0 ? '모든 원본 컬럼이 연결되었습니다.' : '찾는 컬럼이 없습니다.' }));
    }
    shown.forEach((a) => {
      trayList.appendChild(
        el(
          'div',
          { className: `wb-tray-item${st.source.emptyColumns.has(a) ? ' empty' : ''}`, dataset: { dragA: a }, title: '끌어서 타겟 컬럼에 놓으면 연결됩니다' },
          icon('grip', 'wb-grip'),
          el('span', { className: 'wb-card-text one-line' }, el('b', { text: a }), ' ', el('span', { className: 'wb-sample', text: st.source.emptyColumns.has(a) ? '(값 없음)' : sampleText(a, 2) }))
        )
      );
    });
    trayFooter.innerHTML = '';
    if (empty.length > 0) {
      trayFooter.append(
        el('span', { text: ui.showEmptyTray ? `값이 모두 빈 컬럼 ${empty.length}개 포함` : `값이 모두 빈 컬럼 ${empty.length}개 숨김` }),
        el('span', { className: 'wb-spacer' }),
        el('button', {
          type: 'button',
          className: 'wb-link',
          id: 'wbToggleEmpty',
          text: ui.showEmptyTray ? '숨기기' : '보기',
          onclick: () => {
            ui.showEmptyTray = !ui.showEmptyTray;
            renderTray();
          },
        })
      );
    }
  }

  /* ---------- 연결 보기 조작: 클릭 · 키보드 · 끌어다 놓기 ---------- */

  let suppressClickUntil = 0;

  rowsScroller.addEventListener('click', (e) => {
    if (Date.now() < suppressClickUntil) return;
    const toggle = e.target.closest('[data-group-toggle]');
    if (toggle) {
      const g = toggle.dataset.groupToggle;
      ui.collapsed[g] = !ui.collapsed[g];
      renderConnect();
      const again = rowsLayer.querySelector(`[data-item-key="head:${g}"]`);
      if (again) again.focus({ preventScroll: true });
      return;
    }
    const actionBtn = e.target.closest('[data-action]');
    if (actionBtn && actionBtn.dataset.action === 'confirm-all') {
      confirmAllFuzzy();
      return;
    }
    const row = e.target.closest('.wb-row');
    if (!row) return;
    const b = row.dataset.b;
    ui.focusedB = b;
    const cand = e.target.closest('[data-connect-a]');
    if (cand) {
      connectAndFocus(b, cand.dataset.connectA);
      return;
    }
    if (actionBtn) {
      runRowAction(b, actionBtn.dataset.action, actionBtn);
      return;
    }
    if (e.target.closest('[data-pick]')) openPicker(b, row.querySelector('.wb-b-card'));
  });

  rowsScroller.addEventListener('focusin', (e) => {
    const row = e.target.closest('.wb-row');
    if (!row) return;
    const b = row.dataset.b;
    if (ui.focusedB !== b) {
      ui.focusedB = b;
      rowsLayer.querySelectorAll('.wb-row.focused').forEach((r) => r.classList.remove('focused'));
      row.classList.add('focused');
      rowsLayer.querySelectorAll('.wb-row[tabindex="0"]').forEach((r) => r !== row && (r.tabIndex = -1));
      row.tabIndex = 0;
    }
    highlightPreviewColumn(b);
    scrollPreviewToColumn(b, { flash: false });
  });

  rowsScroller.addEventListener('mouseover', (e) => {
    const row = e.target.closest('.wb-row');
    highlightPreviewColumn(row ? row.dataset.b : ui.focusedB);
  });

  async function runRowAction(b, action, anchor) {
    if (action === 'confirm') {
      const next = neighborProblem(b);
      const ok = await api.confirmMatch(b);
      if (ok) focusRow(next || b);
    } else if (action === 'pick') {
      openPicker(b, anchor.closest('.wb-row').querySelector('.wb-b-card'));
    } else if (action === 'disconnect') {
      const ok = await api.disconnect(b);
      if (ok) focusRow(b);
    } else if (action === 'required') {
      await api.toggleRequired(b);
      focusRow(b);
    }
  }

  async function connectAndFocus(b, a) {
    const next = neighborProblem(b);
    const ok = await api.connect(b, a);
    if (ok) focusRow(ui.groupMode === 'status' ? next || b : b);
  }

  async function confirmAllFuzzy() {
    const list = st.headersB.filter((b) => st.matchReason[b] === 'fuzzy');
    if (list.length === 0) return;
    const ok = await showConfirmDialog({
      title: `비슷한 이름 연결 ${list.length}개를 모두 맞음으로 할까요?`,
      message: '아래 연결을 확인한 것으로 표시합니다.',
      items: list.map((b) => `${api.labelOfB(b)} ← ${st.mapping[b]}`),
      note: '실행 취소 한 번으로 모두 되돌릴 수 있습니다.',
      okText: '모두 맞음',
    });
    if (ok) await api.confirmAllFuzzy(list);
  }

  /** 화면 순서상 b 다음(없으면 이전)의 "확인 필요/연결 없음" 줄 — 처리 후 초점을 넘길 곳 */
  function neighborProblem(b) {
    const rows = ui.items.filter((it) => it.type === 'row').map((it) => it.b);
    const problems = rows.filter((x) => x !== b && isProblemReason(st.matchReason[x], st.mapping[x]));
    const idx = rows.indexOf(b);
    return problems.find((x) => rows.indexOf(x) > idx) || problems.filter((x) => rows.indexOf(x) < idx).pop() || null;
  }

  function stepProblem(dir) {
    const rows = ui.items.filter((it) => it.type === 'row').map((it) => it.b);
    const problems = rows.filter((x) => isProblemReason(st.matchReason[x], st.mapping[x]));
    if (problems.length === 0) return;
    const idx = ui.focusedB ? rows.indexOf(ui.focusedB) : -1;
    let target;
    if (dir > 0) target = problems.find((x) => rows.indexOf(x) > idx) || problems[0];
    else target = [...problems].reverse().find((x) => rows.indexOf(x) < idx) || problems[problems.length - 1];
    focusRow(target);
  }

  function stepRow(dir) {
    const rows = ui.items.filter((it) => it.type === 'row').map((it) => it.b);
    if (rows.length === 0) return;
    const idx = ui.focusedB ? rows.indexOf(ui.focusedB) : -1;
    const next = rows[Math.max(0, Math.min(rows.length - 1, idx + dir))];
    focusRow(next);
  }

  /** b 줄로 스크롤하고 초점을 준다(접힌 묶음이면 펼친다) */
  function focusRow(b) {
    if (!b || !st || !st.source || !ui.visual) return;
    if (ui.view !== 'connect') {
      ui.view = 'connect';
      render();
    }
    let idx = ui.items.findIndex((it) => it.type === 'row' && it.b === b);
    if (idx < 0) {
      if (ui.groupMode === 'status') ui.collapsed[groupOf(b)] = false;
      if (!rowMatchesSearch(b)) {
        ui.search = '';
        searchInput.value = '';
      }
      buildItems();
      rowsLayer.style.height = `${Math.max(ui.totalHeight, 1)}px`;
      idx = ui.items.findIndex((it) => it.type === 'row' && it.b === b);
      if (idx < 0) return;
    }
    ui.focusedB = b;
    const y = ui.offsets[idx];
    const height = ui.layout.rowsHeight;
    if (y < ui.layout.rowsTop || y + ROW_H > ui.layout.rowsTop + height) {
      setRowsScrollTop(y - Math.max(0, (height - ROW_H) / 2));
    }
    renderConnectWindow();
    const node = rowsLayer.querySelector(`[data-item-key="${CSS.escape(`row:${b}`)}"]`);
    if (node) node.focus({ preventScroll: true });
  }

  function jumpToProblem(kind) {
    const target = st.headersB.find((b) => (kind === 'fuzzy' ? st.matchReason[b] === 'fuzzy' : !st.mapping[b]));
    if (!target) return;
    if (ui.view !== 'connect') setView('connect');
    focusRow(target);
  }

  rowsScroller.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT') return;
    const row = e.target.closest('.wb-row');
    const onRow = row && e.target === row;
    const code = e.code;
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    // 초점 이벤트가 늦거나 빠진 경우에도 지금 키를 받은 줄을 기준으로 움직인다
    if (row) ui.focusedB = row.dataset.b;
    // 한글 자판에서도 동작하도록 글자(e.key)가 아닌 키 위치(e.code)로 판단한다
    if (code === 'KeyJ' || code === 'KeyK') {
      e.preventDefault();
      stepProblem(code === 'KeyJ' ? 1 : -1);
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      stepRow(e.key === 'ArrowDown' ? 1 : -1);
    } else if (onRow && code === 'KeyY') {
      e.preventDefault();
      if (st.matchReason[row.dataset.b] === 'fuzzy') runRowAction(row.dataset.b, 'confirm');
    } else if (onRow && (code === 'KeyC' || e.key === 'Enter')) {
      e.preventDefault();
      openPicker(row.dataset.b, row.querySelector('.wb-b-card'));
    } else if (onRow && (e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      if (st.mapping[row.dataset.b]) runRowAction(row.dataset.b, 'disconnect');
    }
  });

  /* 끌어다 놓기: 원본 카드(트레이 포함)를 타겟 카드나 빈 자리에 놓는다 */
  let drag = null;
  const DRAG_THRESHOLD = 4;

  function onDragDown(e) {
    if (e.button !== 0) return;
    const src = e.target.closest('[data-drag-a]');
    if (!src || e.target.closest('button')) return;
    drag = { a: src.dataset.dragA, src, x: e.clientX, y: e.clientY, active: false, target: null, overlay: null, line: null };
    window.addEventListener('pointermove', onDragMove);
    window.addEventListener('pointerup', onDragUp);
    window.addEventListener('pointercancel', onDragCancel);
  }

  function onDragMove(e) {
    if (!drag) return;
    if (!drag.active) {
      if (Math.hypot(e.clientX - drag.x, e.clientY - drag.y) < DRAG_THRESHOLD) return;
      drag.active = true;
      closePicker();
      document.body.classList.add('wb-dragging');
      drag.src.classList.add('wb-drag-source');
      const NS = 'http://www.w3.org/2000/svg';
      drag.overlay = document.createElementNS(NS, 'svg');
      drag.overlay.setAttribute('class', 'wb-drag-overlay');
      drag.line = document.createElementNS(NS, 'line');
      drag.overlay.appendChild(drag.line);
      document.body.appendChild(drag.overlay);
    }
    const r = drag.src.getBoundingClientRect();
    drag.line.setAttribute('x1', r.right);
    drag.line.setAttribute('y1', r.top + r.height / 2);
    drag.line.setAttribute('x2', e.clientX);
    drag.line.setAttribute('y2', e.clientY);
    const hit = document.elementFromPoint(e.clientX, e.clientY);
    const target = hit ? hit.closest('[data-drop-b]') : null;
    if (drag.target && drag.target !== target) drag.target.classList.remove('wb-drop-target');
    if (target) target.classList.add('wb-drop-target');
    drag.target = target;
    // 끄는 동안 목록 위·아래 가장자리에 가까우면 자동으로 스크롤
    const sr = rowsScroller.getBoundingClientRect();
    if (e.clientX >= sr.left && e.clientX <= sr.right) {
      if (e.clientY < sr.top + 40) rowsScroller.scrollTop -= 20;
      else if (e.clientY > sr.bottom - 40) rowsScroller.scrollTop += 20;
    }
  }

  function endDrag() {
    window.removeEventListener('pointermove', onDragMove);
    window.removeEventListener('pointerup', onDragUp);
    window.removeEventListener('pointercancel', onDragCancel);
    if (!drag) return null;
    const finished = drag;
    drag = null;
    if (finished.overlay) finished.overlay.remove();
    if (finished.target) finished.target.classList.remove('wb-drop-target');
    finished.src.classList.remove('wb-drag-source');
    document.body.classList.remove('wb-dragging');
    return finished;
  }

  function onDragUp(e) {
    if (drag && drag.active) {
      // 놓은 위치를 마지막으로 다시 확인한다(포인터 이동 없이 바로 놓은 경우)
      const hit = document.elementFromPoint(e.clientX, e.clientY);
      const target = hit ? hit.closest('[data-drop-b]') : drag.target;
      if (target) drag.target = target;
    }
    const finished = endDrag();
    if (!finished || !finished.active) return;
    suppressClickUntil = Date.now() + 300;
    if (finished.target) {
      const b = finished.target.dataset.dropB;
      ui.focusedB = b;
      connectAndFocus(b, finished.a);
    }
  }

  function onDragCancel() {
    endDrag();
  }

  rowsScroller.addEventListener('pointerdown', onDragDown);
  trayList.addEventListener('pointerdown', onDragDown);

  /* ---------- 원본 컬럼 고르기 팝오버 ---------- */

  function openPicker(b, anchorEl) {
    closePicker();
    const label = api.labelOfB(b);
    const input = el('input', { type: 'text', id: 'wbPickerSearch', placeholder: '이름, 초성(ㅎㄷㅍ), 값(010-)으로 찾기', 'aria-label': '원본 컬럼 찾기' });
    const list = el('div', { className: 'wb-picker-list', id: 'wbPickerList', role: 'listbox', 'aria-label': `${label}에 연결할 원본 컬럼` });
    const closeBtn = el('button', { type: 'button', className: 'wb-x', 'aria-label': '닫기', html: ICONS.x });
    const pop = el(
      'div',
      { className: 'wb-picker', id: 'wbPicker', role: 'dialog', 'aria-label': `${label}에 연결할 원본 컬럼 고르기` },
      el('div', { className: 'wb-picker-head' }, el('b', { text: `${label}에 연결할 원본 컬럼` }), el('span', { className: 'wb-muted', text: `${st.headersA.length}열 중` }), el('span', { className: 'wb-spacer' }), closeBtn),
      el('label', { className: 'wb-search focus' }, icon('search'), input),
      list,
      el('div', { className: 'wb-picker-foot', text: '↑↓ 이동 · Enter 연결 · Esc 닫기' })
    );
    document.body.appendChild(pop);
    ui.picker = { b, el: pop, input, list, query: '', active: 0, expanded: { used: false, empty: false }, buttons: [], anchorEl };
    renderPickerList();
    positionPicker();
    input.focus();

    input.addEventListener('input', () => {
      ui.picker.query = input.value;
      ui.picker.active = 0;
      renderPickerList();
    });
    input.addEventListener('keydown', (e) => {
      const p = ui.picker;
      if (!p) return;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        if (p.buttons.length === 0) return;
        p.active = (p.active + (e.key === 'ArrowDown' ? 1 : -1) + p.buttons.length) % p.buttons.length;
        updatePickerActive();
      } else if (e.key === 'Enter') {
        e.preventDefault();
        const btn = p.buttons[p.active];
        if (btn) btn.click();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        closePicker(true);
      }
    });
    closeBtn.addEventListener('click', () => closePicker(true));
    list.addEventListener('mousedown', (e) => e.preventDefault()); // 입력창 초점 유지
    setTimeout(() => document.addEventListener('pointerdown', onOutsidePicker, true), 0);
  }

  function positionPicker() {
    const p = ui.picker;
    if (!p) return;
    const anchor = p.anchorEl && p.anchorEl.isConnected ? p.anchorEl : rowsScroller;
    const rect = anchor.getBoundingClientRect();
    const width = p.el.offsetWidth || 440;
    const height = p.el.offsetHeight || 420;
    const left = Math.max(8, Math.min(rect.right - width, window.innerWidth - width - 8));
    let top = rect.bottom + 6;
    if (top + height > window.innerHeight - 8) top = Math.max(8, rect.top - height - 6);
    p.el.style.left = `${left}px`;
    p.el.style.top = `${top}px`;
  }

  function onOutsidePicker(e) {
    if (ui.picker && !ui.picker.el.contains(e.target)) closePicker();
  }

  function closePicker(returnFocus) {
    if (!ui.picker) return;
    const { b } = ui.picker;
    ui.picker.el.remove();
    ui.picker = null;
    document.removeEventListener('pointerdown', onOutsidePicker, true);
    if (returnFocus) focusRow(b);
  }

  function renderPickerList() {
    const p = ui.picker;
    if (!p || !st || !st.source) return;
    const { b } = p;
    const current = st.mapping[b];
    const counts = usageCount();
    const usedElsewhere = new Map(); // A -> [B 라벨]
    st.headersB.forEach((x) => {
      const a = st.mapping[x];
      if (a && x !== b) {
        if (!usedElsewhere.has(a)) usedElsewhere.set(a, []);
        usedElsewhere.get(a).push(api.labelOfB(x));
      }
    });
    const used = new Set(counts.keys());
    p.list.innerHTML = '';
    p.buttons = [];

    const section = (title, extra) => p.list.appendChild(el('div', { className: 'wb-picker-label' }, el('span', { text: title }), extra ? el('span', { className: 'wb-muted', text: extra }) : null));
    const item = (a, reasons = [], note = '') => {
      const btn = el(
        'button',
        { type: 'button', className: 'wb-pick-item', role: 'option', dataset: { a } },
        el('span', { className: 'wb-card-text one-line' }, el('b', { text: a }), ' ', el('span', { className: 'wb-sample', text: st.source.emptyColumns.has(a) ? '(값 없음)' : sampleText(a, 2) }), note ? el('span', { className: 'wb-muted', text: ` · ${note}` }) : null),
        ...reasons.map((r) => el('span', { className: 'wb-reason', text: r }))
      );
      btn.addEventListener('click', () => choosePicked(b, a));
      p.list.appendChild(btn);
      p.buttons.push(btn);
    };
    const expander = (key, text) => {
      const btn = el('button', { type: 'button', className: 'wb-pick-expand', 'aria-expanded': String(p.expanded[key]) }, icon(p.expanded[key] ? 'chevDown' : 'chevRight'), text);
      btn.addEventListener('click', () => {
        p.expanded[key] = !p.expanded[key];
        const keepIndex = p.buttons.indexOf(btn);
        renderPickerList();
        p.active = Math.max(0, keepIndex);
        updatePickerActive();
      });
      p.list.appendChild(btn);
      p.buttons.push(btn);
    };
    const noteFor = (a) => (a === current ? '지금 연결됨' : usedElsewhere.has(a) ? `${usedElsewhere.get(a).join(', ')}에 연결됨` : '');

    const q = p.query.trim();
    if (q) {
      const found = st.headersA.filter((a) => matchesColumnSearch(q, a, st.source.samples[a]));
      section(`찾은 컬럼 ${found.length}`);
      if (found.length === 0) p.list.appendChild(el('div', { className: 'wb-muted wb-picker-none', text: '맞는 컬럼이 없습니다.' }));
      found.forEach((a) => item(a, [], noteFor(a)));
    } else {
      const recs = recommendFor(b, 3, used);
      const recSet = new Set(recs.map((r) => r.a));
      if (recs.length > 0) {
        section('추천', '이름이 비슷한 정도 + 값 모양');
        recs.forEach((r) => item(r.a, r.reasons));
      }
      const free = st.headersA.filter((a) => !used.has(a) && !st.source.emptyColumns.has(a) && !recSet.has(a));
      section(`연결 안 된 컬럼 ${free.length}`);
      free.forEach((a) => item(a));
      const others = st.headersA.filter((a) => used.has(a));
      if (others.length > 0) {
        expander('used', `이미 연결된 컬럼 ${others.length}`);
        if (p.expanded.used) others.forEach((a) => item(a, [], noteFor(a)));
      }
      const empties = st.headersA.filter((a) => !used.has(a) && st.source.emptyColumns.has(a));
      if (empties.length > 0) {
        expander('empty', `값이 모두 빈 컬럼 ${empties.length}`);
        if (p.expanded.empty) empties.forEach((a) => item(a));
      }
    }
    if (current) {
      const btn = el('button', { type: 'button', className: 'wb-pick-item danger', id: 'wbPickerDisconnect', text: `연결 끊기 (지금: ${current})` });
      btn.addEventListener('click', async () => {
        closePicker();
        const ok = await api.disconnect(b);
        if (ok) focusRow(b);
      });
      p.list.appendChild(btn);
      p.buttons.push(btn);
    }
    if (p.active >= p.buttons.length) p.active = 0;
    updatePickerActive();
  }

  function updatePickerActive() {
    const p = ui.picker;
    if (!p) return;
    p.buttons.forEach((btn, i) => btn.classList.toggle('active', i === p.active));
    const active = p.buttons[p.active];
    if (active) active.scrollIntoView({ block: 'nearest' });
  }

  async function choosePicked(b, a) {
    closePicker();
    if (st.mapping[b] === a && st.matchReason[b] !== 'fuzzy') {
      focusRow(b);
      return;
    }
    await connectAndFocus(b, a);
  }

  /* =====================================================================
   * 목록 보기 (기존 매핑 표를 같은 상태로)
   * ===================================================================== */

  let listRenderToken = 0;

  function renderList() {
    const token = ++listRenderToken;
    const active = document.activeElement;
    const restore = active && root.listBody.contains(active) ? { b: active.dataset.bcolumn, tag: active.tagName } : null;
    root.listBody.innerHTML = '';
    if (!st || !st.source) {
      root.listBody.appendChild(el('tr', {}, el('td', { colspan: '6', className: 'empty-hint', text: '먼저 [구조 분석]을 실행하세요.' })));
      return;
    }
    const lazy = st.headersA.length * st.headersB.length > LIST_LAZY_OPTIONS;
    const fillOptions = (select, current) => {
      select.innerHTML = '';
      select.appendChild(el('option', { value: '', text: '-- 선택 안함 --' }));
      st.headersA.forEach((a) => select.appendChild(el('option', { value: a, text: a })));
      select.value = current || '';
    };
    const buildRow = (b) => {
      const a = st.mapping[b];
      const reason = a ? st.matchReason[b] || 'manual' : 'none';
      const select = el('select', { dataset: { bcolumn: b }, 'aria-label': `${api.labelOfB(b)}의 원본 컬럼` });
      if (lazy) {
        select.appendChild(el('option', { value: '', text: '-- 선택 안함 --' }));
        if (a) select.appendChild(el('option', { value: a, text: a }));
        select.value = a || '';
        const fill = () => {
          if (select.dataset.filled) return;
          select.dataset.filled = '1';
          fillOptions(select, select.value);
        };
        select.addEventListener('focus', fill);
        select.addEventListener('mousedown', fill);
      } else {
        fillOptions(select, a);
      }
      select.addEventListener('change', () => {
        if (select.value) api.connect(b, select.value).then((ok) => !ok && render());
        else api.disconnect(b).then((ok) => !ok && render());
      });

      const checkbox = el('input', { type: 'checkbox', dataset: { bcolumn: b }, 'aria-label': `${api.labelOfB(b)} 필수` });
      checkbox.checked = st.required.has(b);
      checkbox.addEventListener('change', () => api.setRequired(b, checkbox.checked));

      const preview = a
        ? st.source.rows
            .slice(0, 2)
            .map((row) => formatPreviewValue(row[a]))
            .filter((v) => v !== '')
            .join(' / ')
        : '';

      const reasonCell = el('td', {}, el('span', { className: `wb-badge reason-${reason}`, text: REASON_LABELS[reason] }));
      if (reason === 'fuzzy') {
        reasonCell.append(' ', el('button', { type: 'button', className: 'wb-btn warn small', text: '맞음', onclick: () => api.confirmMatch(b) }));
      }
      return el(
          'tr',
          { className: reason === 'fuzzy' ? 'row-fuzzy' : !a ? 'row-none' : '' },
          el('td', { text: b }),
          reasonCell,
          el('td', {}, select),
          el('td', { className: 'mapping-preview-cell', text: preview, title: preview }),
          el('td', { style: 'text-align:center' }, checkbox),
          el('td', {}, a ? el('button', { type: 'button', className: 'wb-x', 'aria-label': `${api.labelOfB(b)} 연결 끊기`, html: ICONS.x, onclick: () => api.disconnect(b) }) : null)
        );
    };
    // 줄이 많으면 앞부분만 바로 그리고 나머지는 잇달아 나눠 그린다(전환이 멈춘 듯 보이지 않게)
    const appendRows = (from) => {
      if (token !== listRenderToken) return;
      const frag = document.createDocumentFragment();
      const to = Math.min(st.headersB.length, from + LIST_BATCH);
      for (let k = from; k < to; k++) frag.appendChild(buildRow(st.headersB[k]));
      root.listBody.appendChild(frag);
      if (to < st.headersB.length) setTimeout(() => appendRows(to), 0);
    };
    appendRows(0);
    if (restore && restore.b) {
      const again = root.listBody.querySelector(`${restore.tag === 'SELECT' ? 'select' : 'input'}[data-bcolumn="${CSS.escape(restore.b)}"]`);
      if (again) again.focus({ preventScroll: true });
    }
  }

  /* =====================================================================
   * 결과 미리보기
   * ===================================================================== */

  function setRowFilter(on) {
    ui.preview.rowFilter = on;
    ui.preview.page = 0;
    if (on) {
      // 켜는 순간의 문제 행을 고정해 둔다 — 고쳐서 문제가 풀린 행도 "해결"로 남아 보이게
      const derived = api.getDerived();
      const editedRows = new Set(Array.from(st.overrides.keys()).map((k) => Number(k.slice(0, k.indexOf('|')))));
      const list = [];
      for (let i = 0; i < derived.rowCount; i++) {
        if (derived.missingByRow[i] || derived.duplicateOf[i] >= 0 || editedRows.has(i)) list.push(i);
      }
      ui.preview.rowFilterList = list;
    } else {
      ui.preview.rowFilterList = null;
    }
    renderPreview();
  }

  function previewColumns() {
    const derived = api.getDerived();
    const problem = (b) => ['reqmissing', 'fuzzy', 'none'].includes(derived.columnStatus[b]);
    const visible = ui.preview.colFilter ? st.headersB.filter((b) => ui.preview.pinned.has(b) || problem(b)) : st.headersB;
    return {
      pinned: visible.filter((b) => ui.preview.pinned.has(b)),
      scroll: visible.filter((b) => !ui.preview.pinned.has(b)),
      problemCount: st.headersB.filter(problem).length,
    };
  }

  /** 가로로 보이는 열 범위. 열이 적으면 전부 */
  function computeColumnRange() {
    const { scroll, pinned } = previewColumns();
    if (scroll.length <= H_VIRTUAL_THRESHOLD) return { first: 0, last: scroll.length - 1, virtual: false };
    const pinnedWidth = ROWNUM_W + pinned.length * COL_W;
    const viewWidth = Math.max(COL_W, ui.layout.pvWidth - pinnedWidth);
    const first = Math.max(0, Math.floor(ui.layout.pvLeft / COL_W) - 2);
    const last = Math.min(scroll.length - 1, first + Math.ceil(viewWidth / COL_W) + 4);
    return { first, last, virtual: true };
  }

  function renderPreview() {
    if (!st || !st.source || !ui.visual) return;
    if (ui.preview.editing) return; // 입력 중에는 다시 그리지 않는다(입력 내용 보존)
    const derived = api.getDerived();
    const n = derived.rowCount;
    const cols = previewColumns();
    const showColTools = st.headersB.length > PIN_THRESHOLD;
    pvColSearchWrap.hidden = !showColTools;
    pvColSeg.hidden = !showColTools;
    pvColAll.setAttribute('aria-pressed', String(!ui.preview.colFilter));
    pvColProblem.setAttribute('aria-pressed', String(ui.preview.colFilter));
    pvColProblem.textContent = `문제 있는 열만 ${cols.problemCount}`;
    pvRowFilter.setAttribute('aria-pressed', String(ui.preview.rowFilter));
    pvRowFilter.classList.toggle('on', ui.preview.rowFilter);

    // 행
    const indices = ui.preview.rowFilter ? ui.preview.rowFilterList || [] : null;
    const totalRows = indices ? indices.length : n;
    const pages = Math.max(1, Math.ceil(totalRows / PREVIEW_PAGE));
    ui.preview.page = Math.min(ui.preview.page, pages - 1);
    const start = ui.preview.page * PREVIEW_PAGE;
    const end = Math.min(totalRows, start + PREVIEW_PAGE);
    const pageRows = [];
    for (let k = start; k < end; k++) pageRows.push(indices ? indices[k] : k);
    pvPrev.disabled = ui.preview.page === 0;
    pvNext.disabled = ui.preview.page >= pages - 1;

    // 열
    const range = computeColumnRange();
    ui.preview.range = range;
    const shownScroll = cols.scroll.slice(range.first, range.last + 1);
    const leftSpace = range.first * COL_W;
    const rightSpace = (cols.scroll.length - 1 - range.last) * COL_W;
    const tableWidth = ROWNUM_W + (cols.pinned.length + cols.scroll.length) * COL_W;
    pvTable.style.width = `${tableWidth}px`;

    pvRange.textContent =
      `${ui.preview.rowFilter ? `문제 있는 행 ${totalRows.toLocaleString()}개` : `전체 ${n.toLocaleString()}행`} 중 ${totalRows ? `${start + 1}–${end}` : '0'}` +
      (range.virtual ? ` · ${st.headersB.length}열 중 ${range.first + 1 + cols.pinned.length}–${range.last + 1 + cols.pinned.length}열` : '');

    const active = ui.preview.active;

    // 머리글: 열 구성과 열 상태가 그대로면 다시 만들지 않는다
    const headSig = [cols.pinned.join('|'), shownScroll.join('|'), leftSpace, rightSpace, ...[...cols.pinned, ...shownScroll].map((b) => headerKey(b, derived))].join('#');
    if (headSig !== ui.preview.headSig) {
      ui.preview.headSig = headSig;
      pvHead.innerHTML = '';
      ui.preview.thByB = new Map();
      const headRow = el('tr', {}, el('th', { className: 'wb-rownum sticky', style: `left:0;width:${ROWNUM_W}px`, text: '원본 행' }));
      cols.pinned.forEach((b, k) => headRow.appendChild(buildHeaderCell(b, derived, { left: ROWNUM_W + k * COL_W, edge: k === cols.pinned.length - 1 })));
      if (leftSpace > 0) headRow.appendChild(el('th', { className: 'wb-spacer-cell', style: `width:${leftSpace}px` }));
      shownScroll.forEach((b) => headRow.appendChild(buildHeaderCell(b, derived, null)));
      if (rightSpace > 0) headRow.appendChild(el('th', { className: 'wb-spacer-cell', style: `width:${rightSpace}px` }));
      pvHead.appendChild(headRow);
    }

    // 본문: 보이는 행·열이 지난번과 같으면 칸을 새로 만들지 않고 바뀐 칸만 고친다
    // (연결 하나·값 하나를 바꿀 때마다 표 전체를 다시 만들면 레이아웃 비용이 커진다)
    const rowNumbers = st.source.rowNumbers;
    const describeCell = (i, b, missing, pinnedLeft, edge) => {
      const v = derived.value(i, b);
      const isActive = active && active.i === i && active.b === b;
      const edited = api.getOverride(i, b) !== undefined;
      const isMissing = !!missing && missing.includes(b);
      const classes = ['wb-cell'];
      const align = api.cellAlign(v, b);
      if (align) classes.push(align);
      if (isMissing) classes.push('missing');
      if (edited) classes.push('edited');
      if (pinnedLeft !== null) classes.push('sticky');
      if (edge) classes.push('pin-edge');
      if (isActive) classes.push('active');
      return {
        className: classes.join(' '),
        text: api.displayValue(v, b),
        title: isMissing ? '필수 칸이 비어 있음' : edited ? '직접 고친 값' : '',
        tabIndex: isActive ? 0 : -1,
      };
    };
    const rowNumKey = (i, missing) => {
      const dupOf = derived.duplicateOf[i];
      const mark = dupOf >= 0 ? `dup:${dupOf}` : ui.preview.rowFilter && !missing ? 'resolved' : '';
      return `${rowNumbers[i]}|${mark}|${missing ? 1 : 0}`;
    };
    const fillRowNum = (td, i, missing) => {
      const dupOf = derived.duplicateOf[i];
      td.className = 'wb-rownum sticky';
      td.textContent = String(rowNumbers[i]);
      td.title = '';
      if (dupOf >= 0) {
        td.classList.add('dup');
        td.title = `${rowNumbers[dupOf]}행과 모든 값이 같음`;
        td.append(el('span', { className: 'wb-dup', text: '중복' }));
      } else if (ui.preview.rowFilter && !missing) {
        // 필터를 켤 때는 문제였지만 지금은 빈칸·중복이 없는 행
        td.append(el('span', { className: 'wb-resolved', text: '해결' }));
      }
      if (missing) td.classList.add('missing');
      td.dataset.key = rowNumKey(i, missing);
    };
    const columnsOfRow = [
      ...cols.pinned.map((b, k) => [b, ROWNUM_W + k * COL_W, k === cols.pinned.length - 1]),
      ...shownScroll.map((b) => [b, null, false]),
    ];
    const bodySig = [pageRows.join(','), cols.pinned.join('|'), shownScroll.join('|'), leftSpace, rightSpace].join('#');
    if (bodySig === ui.preview.bodySig && pvBody.rows.length === pageRows.length) {
      pageRows.forEach((i, ri) => {
        const tr = pvBody.rows[ri];
        const missing = derived.missingByRow[i];
        const rowNumCell = tr.cells[0];
        if (rowNumCell.dataset.key !== rowNumKey(i, missing)) fillRowNum(rowNumCell, i, missing);
        const cells = tr.querySelectorAll('td.wb-cell');
        columnsOfRow.forEach(([b, pinnedLeft, edge], ci) => {
          const td = cells[ci];
          const want = describeCell(i, b, missing, pinnedLeft, edge);
          if (td.className !== want.className) td.className = want.className;
          if (td.textContent !== want.text || td.childElementCount > 0) td.textContent = want.text;
          if (td.title !== want.title) td.title = want.title;
          if (td.tabIndex !== want.tabIndex) td.tabIndex = want.tabIndex;
        });
      });
    } else {
      ui.preview.bodySig = bodySig;
      pvBody.innerHTML = '';
      pageRows.forEach((i) => {
        const missing = derived.missingByRow[i];
        const rowNumCell = el('td', { style: 'left:0' });
        fillRowNum(rowNumCell, i, missing);
        const tr = el('tr', { dataset: { r: String(i) } }, rowNumCell);
        const addCell = ([b, pinnedLeft, edge]) => {
          const want = describeCell(i, b, missing, pinnedLeft, edge);
          tr.appendChild(
            el('td', {
              className: want.className,
              tabindex: String(want.tabIndex),
              style: pinnedLeft !== null ? `left:${pinnedLeft}px` : null,
              dataset: { r: String(i), b },
              title: want.title || null,
              text: want.text,
            })
          );
        };
        columnsOfRow.slice(0, cols.pinned.length).forEach(addCell);
        if (leftSpace > 0) tr.appendChild(el('td', { className: 'wb-spacer-cell' }));
        columnsOfRow.slice(cols.pinned.length).forEach(addCell);
        if (rightSpace > 0) tr.appendChild(el('td', { className: 'wb-spacer-cell' }));
        pvBody.appendChild(tr);
      });
    }
    // 키보드로 표에 들어올 수 있도록 활성 칸이 없으면 첫 칸을 탭 정지점으로
    if (!pvBody.querySelector('td[tabindex="0"]')) {
      const first = pvBody.querySelector('td.wb-cell');
      if (first) first.tabIndex = 0;
    }

    pvEmpty.hidden = pageRows.length > 0;
    pvEmpty.textContent = ui.preview.rowFilter ? '문제 있는 행이 없습니다.' : '변환할 데이터 행이 없습니다.';
    pvFooterNote.textContent = cols.pinned.length
      ? `고정: 원본 행 · ${cols.pinned.map((b) => api.labelOfB(b)).join(' · ')} (열 머리글의 핀으로 바꿈)`
      : '열 머리글의 핀을 누르면 그 열을 왼쪽에 고정합니다';

    renderColMap(derived, cols);
    highlightPreviewColumn(ui.focusedB);
  }

  /** 머리글 칸 하나의 모양을 정하는 값들 — 같으면 다시 그리지 않는다 */
  function headerKey(b, derived) {
    return [api.labelOfB(b), st.mapping[b] || '', st.matchReason[b] || '', derived.columnStatus[b], st.required.has(b), ui.preview.pinned.has(b)].join('|');
  }

  function buildHeaderCell(b, derived, pin) {
    const a = st.mapping[b];
    const reason = a ? st.matchReason[b] || 'manual' : 'none';
    const status = derived.columnStatus[b];
    const required = st.required.has(b);
    const pinned = ui.preview.pinned.has(b);
    const th = el(
      'th',
      {
        className: `wb-col status-${status}${pin ? ' sticky' : ''}${pin && pin.edge ? ' pin-edge' : ''}`,
        style: `width:${COL_W}px${pin ? `;left:${pin.left}px` : ''}`,
        dataset: { b },
        title: '눌러서 연결 보기에서 이 컬럼으로 이동',
      },
      el(
        'div',
        { className: 'wb-th-top' },
        el('span', { className: 'wb-th-label', text: api.labelOfB(b) }),
        required ? el('span', { className: 'wb-req-mark', text: '필수' }) : null,
        el('span', { className: 'wb-colno', text: String(st.headersB.indexOf(b) + 1) }),
        el('span', { className: 'wb-spacer' }),
        el('button', {
          type: 'button',
          className: `wb-pin${pinned ? ' on' : ''}`,
          dataset: { pin: b },
          'aria-pressed': String(pinned),
          'aria-label': `${api.labelOfB(b)} 열 ${pinned ? '고정 풀기' : '왼쪽에 고정'}`,
          title: pinned ? '고정 풀기' : '왼쪽에 고정',
          html: ICONS.pin,
        })
      ),
      el('span', { className: `wb-th-src reason-${reason}`, text: a ? `← ${a}${reason === 'fuzzy' ? ' · 확인 필요' : ''}` : '연결 없음' })
    );
    ui.preview.thByB.set(b, th);
    return th;
  }

  function renderColMap(derived, cols) {
    pvColMap.innerHTML = '';
    pvColMap.hidden = st.headersB.length <= GROUP_THRESHOLD;
    if (pvColMap.hidden) return;
    const legend = (cls, text) => el('span', { className: 'wb-legend' }, el('span', { className: `wb-swatch status-${cls}` }), text);
    pvColMap.append(
      el(
        'div',
        { className: 'wb-colmap-head' },
        el('b', { text: '열 지도' }),
        el('span', { className: 'wb-muted', text: '칸을 누르면 그 열로 이동 · 굵은 테두리는 고정한 열' }),
        el('span', { className: 'wb-spacer' }),
        legend('ok', '정상'),
        legend('fuzzy', '확인 필요'),
        legend('reqmissing', '필수 빈칸'),
        legend('none', '연결 없음'),
        legend('edited', '고친 값')
      )
    );
    const strip = el('div', { className: 'wb-colmap-strip' });
    st.headersB.forEach((b, idx) => {
      strip.appendChild(
        el('button', {
          type: 'button',
          className: `wb-colmap-cell status-${derived.columnStatus[b]}${ui.preview.pinned.has(b) ? ' pinned' : ''}`,
          dataset: { jump: b },
          title: `${idx + 1}열 · ${api.labelOfB(b)}`,
          'aria-label': `${idx + 1}열 ${api.labelOfB(b)}로 이동`,
        })
      );
    });
    pvColMap.appendChild(strip);
    updateColMapViewport(cols);
  }

  /** 열 지도에서 지금 화면에 보이는 열을 표시한다 */
  function updateColMapViewport(colsArg) {
    if (pvColMap.hidden || !st) return;
    const cols = colsArg || previewColumns();
    const pinnedWidth = ROWNUM_W + cols.pinned.length * COL_W;
    const viewWidth = Math.max(COL_W, ui.layout.pvWidth - pinnedWidth);
    const firstVisible = Math.floor(ui.layout.pvLeft / COL_W);
    const lastVisible = Math.min(cols.scroll.length - 1, Math.floor((ui.layout.pvLeft + viewWidth - 1) / COL_W));
    const visible = new Set(cols.pinned);
    for (let k = firstVisible; k <= lastVisible; k++) if (cols.scroll[k]) visible.add(cols.scroll[k]);
    pvColMap.querySelectorAll('.wb-colmap-cell').forEach((cell) => cell.classList.toggle('in-view', visible.has(cell.dataset.jump)));
  }

  pvColMap.addEventListener('click', (e) => {
    const cell = e.target.closest('[data-jump]');
    if (!cell) return;
    const b = cell.dataset.jump;
    if (ui.preview.colFilter && !previewColumns().scroll.includes(b) && !ui.preview.pinned.has(b)) {
      ui.preview.colFilter = false;
      renderPreview();
    }
    scrollPreviewToColumn(b, { flash: true });
  });

  function scrollPreviewToColumn(b, { flash }) {
    if (!st || !st.source || ui.preview.editing) return;
    const cols = previewColumns();
    const k = cols.scroll.indexOf(b);
    if (k >= 0) {
      const pinnedWidth = ROWNUM_W + cols.pinned.length * COL_W;
      const viewWidth = Math.max(COL_W, ui.layout.pvWidth - pinnedWidth);
      const x = k * COL_W;
      if (x < ui.layout.pvLeft || x + COL_W > ui.layout.pvLeft + viewWidth) {
        setPreviewScrollLeft(x - Math.max(0, (viewWidth - COL_W) / 2));
        const range = computeColumnRange();
        const before = ui.preview.range;
        if (!before || before.first !== range.first || before.last !== range.last) renderPreview();
        else updateColMapViewport(cols);
      }
    }
    if (flash) {
      const th = ui.preview.thByB.get(b);
      if (th) {
        th.classList.remove('wb-flash');
        void th.offsetWidth;
        th.classList.add('wb-flash');
      }
    }
  }

  function highlightPreviewColumn(b) {
    ui.preview.thByB.forEach((th, key) => th.classList.toggle('wb-hl', key === b));
  }

  pvHead.addEventListener('click', (e) => {
    const pinBtn = e.target.closest('[data-pin]');
    if (pinBtn) {
      const b = pinBtn.dataset.pin;
      if (ui.preview.pinned.has(b)) ui.preview.pinned.delete(b);
      else ui.preview.pinned.add(b);
      renderPreview();
      return;
    }
    const th = e.target.closest('th[data-b]');
    if (th) focusRow(th.dataset.b);
  });
  pvHead.addEventListener('mouseover', (e) => {
    const th = e.target.closest('th[data-b]');
    rowsLayer.querySelectorAll('.wb-row.hl').forEach((r) => r.classList.remove('hl'));
    if (th) {
      const row = rowsLayer.querySelector(`.wb-row[data-b="${CSS.escape(th.dataset.b)}"]`);
      if (row) row.classList.add('hl');
    }
  });

  /* ---------- 미리보기 칸 조작: 초점 이동 · 고치기 ---------- */

  function cellAt(i, b) {
    return pvBody.querySelector(`td.wb-cell[data-r="${i}"][data-b="${CSS.escape(b)}"]`);
  }

  function setActive(i, b, td) {
    ui.preview.active = { i, b };
    pvBody.querySelectorAll('td.wb-cell.active').forEach((c) => {
      c.classList.remove('active');
      c.tabIndex = -1;
    });
    pvBody.querySelectorAll('td.wb-cell[tabindex="0"]').forEach((c) => (c.tabIndex = -1));
    if (td) {
      td.classList.add('active');
      td.tabIndex = 0;
    }
  }

  function focusCell(i, b) {
    let td = cellAt(i, b);
    if (!td) {
      scrollPreviewToColumn(b, { flash: false });
      td = cellAt(i, b);
    }
    if (!td) return;
    setActive(i, b, td);
    td.focus({ preventScroll: false });
    if (td.classList.contains('edited')) showBubble(td, i, b);
    else hideBubble();
  }

  pvBody.addEventListener('click', (e) => {
    const td = e.target.closest('td.wb-cell');
    if (!td || ui.preview.editing) return;
    focusCell(Number(td.dataset.r), td.dataset.b);
  });
  pvBody.addEventListener('dblclick', (e) => {
    const td = e.target.closest('td.wb-cell');
    if (!td) return;
    startEdit(Number(td.dataset.r), td.dataset.b);
  });
  pvBody.addEventListener('focusin', (e) => {
    const td = e.target.closest('td.wb-cell');
    if (!td || ui.preview.editing) return;
    const i = Number(td.dataset.r);
    const b = td.dataset.b;
    if (!ui.preview.active || ui.preview.active.i !== i || ui.preview.active.b !== b) setActive(i, b, td);
    highlightPreviewColumn(b);
  });
  pvBody.addEventListener('keydown', (e) => {
    const td = e.target.closest('td.wb-cell');
    if (!td || e.target.tagName === 'INPUT') return;
    const i = Number(td.dataset.r);
    const b = td.dataset.b;
    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      startEdit(i, b);
    } else if (e.key.startsWith('Arrow')) {
      e.preventDefault();
      moveActive(i, b, e.key);
    } else if (e.key === 'Escape') {
      hideBubble();
    }
  });

  function visibleColumnOrder() {
    const cols = previewColumns();
    return [...cols.pinned, ...cols.scroll];
  }

  function moveActive(i, b, key) {
    const order = visibleColumnOrder();
    const rowsOnPage = Array.from(pvBody.querySelectorAll('tr[data-r]')).map((tr) => Number(tr.dataset.r));
    let ci = order.indexOf(b);
    let ri = rowsOnPage.indexOf(i);
    if (key === 'ArrowRight') ci = Math.min(order.length - 1, ci + 1);
    if (key === 'ArrowLeft') ci = Math.max(0, ci - 1);
    if (key === 'ArrowDown') ri = Math.min(rowsOnPage.length - 1, ri + 1);
    if (key === 'ArrowUp') ri = Math.max(0, ri - 1);
    focusCell(rowsOnPage[ri], order[ci]);
  }

  function startEdit(i, b) {
    if (ui.preview.editing) return;
    hideBubble();
    let td = cellAt(i, b);
    if (!td) return;
    const override = api.getOverride(i, b);
    const startText = override !== undefined ? override : formatPreviewValue(api.originalValue(i, b));
    ui.preview.editing = { i, b, startText };
    setActive(i, b, td);
    td.classList.add('editing');
    td.textContent = '';
    const input = el('input', { type: 'text', className: 'wb-cell-input', id: 'wbCellInput', value: startText, 'aria-label': `${st.source.rowNumbers[i]}행 ${api.labelOfB(b)}` });
    input.value = startText;
    const hint = el('div', { className: 'wb-edit-hint', text: 'Enter 확정 · Esc 취소 · Tab 다음 칸' });
    td.append(input, hint);
    input.focus();
    input.select();
    let done = false;
    const finish = async (save, move) => {
      if (done) return;
      done = true;
      const value = input.value;
      ui.preview.editing = null;
      let changed = false;
      if (save && value !== startText) {
        changed = await api.setOverride(i, b, value);
        if (changed) {
          const now = api.getOverride(i, b);
          showToast(`${st.source.rowNumbers[i]}행 ${api.labelOfB(b)}을(를) ${now === undefined ? '원래 값으로 되돌렸습니다' : '고쳤습니다'}`);
        }
      }
      if (!changed) renderPreview();
      if (move) {
        const order = visibleColumnOrder();
        const ci = order.indexOf(b) + move;
        if (ci >= 0 && ci < order.length) {
          focusCell(i, order[ci]);
          return;
        }
      }
      focusCell(i, b);
    };
    input.addEventListener('keydown', (e) => {
      if (e.isComposing) return; // 한글 조합 중 Enter는 글자 확정용
      if (e.key === 'Enter') {
        e.preventDefault();
        finish(true, 0);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        finish(false, 0);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        finish(true, e.shiftKey ? -1 : 1);
      }
    });
    input.addEventListener('blur', () => finish(true, 0));
  }

  /* ---------- 단축키: 실행 취소 / 다시 실행 ---------- */

  document.addEventListener('keydown', (e) => {
    if (!st || !st.source) return;
    const tab1 = document.getElementById('tab1');
    if (!tab1 || !tab1.classList.contains('active')) return;
    if (document.querySelector('.modal-overlay:not([hidden])')) return;
    const t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    if (e.code === 'KeyZ') {
      e.preventDefault();
      if (e.shiftKey) api.redo();
      else api.undo();
    } else if (e.code === 'KeyY') {
      e.preventDefault();
      api.redo();
    }
  });

  // 크기 변화는 여기서만 읽는다(숨겨진 동안의 0은 무시)
  const sizeObserver = new ResizeObserver((entries) => {
    let rowsChanged = false;
    let pvChanged = false;
    entries.forEach((entry) => {
      const { height, width } = entry.contentRect;
      if (entry.target === rowsScroller && height > 0 && Math.abs(height - ui.layout.rowsHeight) > 1) {
        ui.layout.rowsHeight = rowsScroller.clientHeight;
        rowsChanged = true;
      }
      if (entry.target === pvWrap && width > 0 && Math.abs(width - ui.layout.pvWidth) > 1) {
        ui.layout.pvWidth = pvWrap.clientWidth;
        pvChanged = true;
      }
    });
    if (!st || !st.source) return;
    if (rowsChanged && ui.view === 'connect') renderConnectWindow();
    if (pvChanged && !ui.preview.editing) renderPreview();
  });
  sizeObserver.observe(rowsScroller);
  sizeObserver.observe(pvWrap);

  window.addEventListener('resize', () => {
    if (!st || !st.source) return;
    if (ui.view === 'connect') renderConnectWindow();
    if (!ui.preview.editing) renderPreview();
    positionPicker();
  });
  document.addEventListener(
    'scroll',
    () => {
      hideBubble();
      if (ui.picker) positionPicker();
    },
    true
  );

  /* ---------- 외부에서 부르는 함수 ---------- */

  function resetUi() {
    closePicker();
    hideBubble();
    toast.hidden = true;
    ui.view = ui.visual ? 'connect' : 'list';
    ui.groupMode = 'order';
    ui.collapsed = { auto: true };
    ui.search = '';
    searchInput.value = '';
    ui.traySearch = '';
    traySearch.value = '';
    ui.showEmptyTray = false;
    ui.focusedB = null;
    ui.recCache = new Map();
    ui.preview.page = 0;
    ui.preview.rowFilter = false;
    ui.preview.rowFilterList = null;
    ui.preview.colFilter = false;
    ui.preview.pinned = new Set();
    ui.preview.active = null;
    ui.preview.editing = null;
    ui.preview.range = null;
    ui.preview.headSig = null;
    ui.preview.bodySig = null;
    pvColSearch.value = '';
    ui.layout.rowsTop = 0;
    rowsScroller.scrollTop = 0;
    setPreviewScrollLeft(0);
    pvWrap.scrollTop = 0;
  }

  return {
    render,
    reset: resetUi,
    onAnalyzed() {
      resetUi();
      const s = api.getState();
      ui.groupMode = s.headersB.length > GROUP_THRESHOLD ? 'status' : 'order';
      if (s.headersB.length > PIN_THRESHOLD && s.headersB[0]) ui.preview.pinned.add(s.headersB[0]);
    },
    onRender(cb) {
      ui.renderCallbacks.push(cb);
    },
    isVisual() {
      return ui.visual;
    },
  };
}

module.exports = { createWorkbench };
