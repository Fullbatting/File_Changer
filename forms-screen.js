// forms-screen.js
// ② 서식 관리 화면: 기관별 서식 목록, 서식 파일·폴더 추가(파일 이름으로 분류 → 확인 창에서 고쳐 등록),
// 서식 검토(필드 연결), 한글 변환, 시험 저장(가짜 값), 삭제, 백업 내보내기·가져오기.

const { ipcRenderer, webUtils } = require('electron');
const fs = require('fs');
const path = require('path');
const { h, replace, select } = require('./ui-dom.js');
const { KINDS, APPLICANTS, COMMON_AGENCY, prepareImport, entryStatus, sortForManager } = require('./form-library.js');
const { fieldOf, sourceLabel, SAMPLE_INPUT } = require('./merge-fields.js');
const { ensureHwpx, saveFilled } = require('./print-job.js');
const { log, notifyError } = require('./common.js');

const FORM_FILTERS = [
  { name: '한글 서식(.hwp, .hwpx)', extensions: ['hwp', 'hwpx', 'hwpt'] },
  { name: 'All Files', extensions: ['*'] },
];
const BACKUP_FILTERS = [{ name: '서식 백업', extensions: ['zip'] }];
const MAPPING_TYPES = [
  ['fixed', '고정 글자'],
  ['ask', '출력할 때 입력'],
  ['blank', '비워 둠'],
];

function init(app) {
  const $ = (id) => document.getElementById(id);
  const state = { agency: '' }; // '' = 전체
  const hancom = (a, b) => app.hancom(a, b);
  const changed = () => app.changed('library'); // 이 화면도 듣고 있어 다시 그려진다
  const fillAgencyList = (names) => replace($('agencyList'), names.map((a) => h('option', { value: a })));
  const docLine = (d) => `${d.mark} ${d.title}`.trim();

  /* ---------- 목록 ---------- */

  function sideButton(name, label, count) {
    const onclick = () => {
      state.agency = name;
      render();
    };
    return h('button', { type: 'button', 'aria-current': String(state.agency === name), onclick }, h('span', {}, label), h('span', {}, String(count)));
  }

  function actionButton(label, onclick, danger) {
    return h('button', { type: 'button', class: `fa-btn small${danger ? ' danger' : ''}`, onclick }, label);
  }

  function entryRow(e) {
    const st = entryStatus(e);
    const places = e.fields.reduce((n, f) => n + f.count, 0);
    const actions = [actionButton('검토', () => openReview(e.id)), st.needsConvert ? actionButton('한글 변환', () => convert([e.id])) : null, actionButton('시험 저장', () => testSave(e.id)), actionButton('삭제', () => removeEntry(e.id), true)].filter(Boolean);
    return h(
      'tr',
      { 'data-id': e.id },
      h('td', {}, h('div', { style: 'font-weight: 700' }, `${e.favorite ? '★ ' : ''}${e.title}`), h('div', { class: 'sub' }, `${state.agency ? '' : `${e.agency} · `}${e.sourceName}`)),
      h('td', { class: 'nowrap' }, e.kind),
      h('td', { class: 'nowrap' }, e.applicant),
      h('td', { class: 'nowrap', title: e.documents.map(docLine).join('\n') }, `${e.documents.length}종`),
      h('td', { class: 'nowrap' }, e.fields.length ? `${e.fields.length}개 · ${places}곳` : '없음(자동 인식)'),
      h('td', { class: 'nowrap' }, h('span', { class: `fa-badge${st.label === '사용 중' ? '' : ' warn'}`, title: st.problems.join('\n') }, st.label)),
      h('td', { class: 'actions' }, actions.flatMap((b, i) => (i ? [' ', b] : [b])))
    );
  }

  function render() {
    const agencies = app.library.agencies();
    if (state.agency && !agencies.some((a) => a.name === state.agency)) state.agency = '';
    const all = app.library.list();
    replace($('formsSide'), sideButton('', '전체', all.length), agencies.map((a) => sideButton(a.name, a.name, a.count)));
    $('formsTitle').textContent = state.agency ? `${state.agency} 서식` : '전체 서식';
    const list = sortForManager(all.filter((e) => !state.agency || e.agency === state.agency));
    replace(
      $('formsTableBody'),
      list.length === 0 ? h('tr', {}, h('td', { colspan: '7' }, h('div', { class: 'fa-empty' }, '등록한 서식이 없습니다. 위의 [서식 파일 추가]·[폴더째 추가]를 누르거나 파일을 끌어다 놓으세요.'))) : list.map(entryRow)
    );
  }

  /* ---------- 추가: 파일 이름으로 분류 → 확인 창 ---------- */

  let pending = []; // 확인 창의 줄들(form-library.js의 prepareImport)

  function startImport(paths) {
    pending = prepareImport(paths, app.library.list());
    if (pending.length === 0) {
      alert('한글 서식 파일(.hwp, .hwpx)을 찾지 못했습니다.');
      return;
    }
    renderImport();
    $('importModal').hidden = false;
  }

  function importRow(p) {
    const name = path.basename(p.file);
    const note = p.error ? h('div', { class: 'sub', style: 'color: var(--fa-danger)' }, `읽을 수 없음: ${p.error}`) : p.dup ? h('div', { class: 'sub' }, '이미 등록한 서식입니다(고르면 하나 더 등록)') : null;
    return h(
      'tr',
      {},
      h('td', { class: 'file' }, h('label', {}, h('input', { type: 'checkbox', checked: p.include, disabled: !!p.error, 'aria-label': `${name} 등록`, onchange: (ev) => (p.include = ev.target.checked) }), ' ', name), note),
      h('td', {}, h('input', { value: p.agency, list: 'agencyList', 'aria-label': '제출 기관', oninput: (ev) => (p.agency = ev.target.value.trim() || COMMON_AGENCY) })),
      h('td', {}, select(KINDS, p.kind, (v) => (p.kind = v), { 'aria-label': '지원 종류' })),
      h('td', {}, select(APPLICANTS, p.applicant, (v) => (p.applicant = v), { 'aria-label': '구분' })),
      h('td', {}, h('input', { type: 'checkbox', checked: p.favorite, 'aria-label': '자주 쓰는 서식', onchange: (ev) => (p.favorite = ev.target.checked) })),
      h('td', {}, h('input', { value: p.title, 'aria-label': '서식 이름', oninput: (ev) => (p.title = ev.target.value.trim() || p.title) })),
      h('td', {}, `필드 ${p.fields.length}개 · 서류 ${p.documents.length}종`, p.unknown.length ? h('div', { class: 'sub', style: 'color: var(--fa-warn)' }, `사전에 없는 필드: ${p.unknown.join(', ')}`) : null)
    );
  }

  function renderImport() {
    fillAgencyList([...new Set([...app.library.agencies().map((a) => a.name), ...pending.map((p) => p.agency), COMMON_AGENCY])]);
    replace($('importTableBody'), pending.map(importRow));
  }

  function closeImport() {
    $('importModal').hidden = true;
    pending = [];
  }

  $('importCancel').addEventListener('click', closeImport);
  $('importConfirm').addEventListener('click', async () => {
    const chosen = pending.filter((p) => p.include);
    closeImport();
    const added = [];
    chosen.forEach((p) => {
      try {
        added.push(app.library.add(p.file, { agency: p.agency, kind: p.kind, applicant: p.applicant, favorite: p.favorite, title: p.title, fields: p.fields, documents: p.documents }));
      } catch (err) {
        notifyError(`${path.basename(p.file)} 등록 중 오류가 발생했습니다`, err);
      }
    });
    log(`서식 ${added.length}개를 등록했습니다.`, 'ok');
    changed();
    const toConvert = added.filter((e) => !e.hwpx).map((e) => e.id);
    if (toConvert.length === 0) return;
    if (process.platform === 'win32') await convert(toConvert);
    else log(`.hwp 서식 ${toConvert.length}개는 한글이 있는 Windows PC에서 처음 출력할 때(또는 [한글 변환]으로) .hwpx로 바꿉니다.`, 'warn');
  });

  $('btnAddFiles').addEventListener('click', async () => {
    const paths = await ipcRenderer.invoke('dialog:openFile', { title: '한글 서식 파일 선택', filters: FORM_FILTERS, multi: true });
    if (paths && paths.length) startImport(paths);
  });
  $('btnAddFolder').addEventListener('click', async () => {
    const dir = await ipcRenderer.invoke('dialog:selectFolder', { title: '서식이 든 폴더 선택' });
    if (dir) startImport([dir]);
  });

  const drop = $('formsDrop');
  drop.addEventListener('dragover', (e) => {
    e.preventDefault();
    drop.classList.add('dragover');
  });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('dragover');
    const paths = [...(e.dataTransfer.files || [])].map((f) => webUtils.getPathForFile(f)).filter(Boolean);
    if (paths.length) startImport(paths);
  });

  /* ---------- 한글 변환 ---------- */

  async function convert(ids) {
    for (const id of ids) {
      const e = app.library.get(id);
      log(`한글로 .hwpx 변환 중: ${e.title}`);
      try {
        await ensureHwpx(app.library, e, { hancom });
        log(`변환 완료: ${e.title}`, 'ok');
      } catch (err) {
        log(`변환 실패: ${e.title} — ${err.message}`, 'warn');
      }
    }
    changed();
  }

  /* ---------- 검토: 이름·기관·종류·구분과 필드 연결 ---------- */

  let reviewing = null; // { id, mapping, meta }

  function metaFields(m) {
    const field = (label, control) => h('div', { class: 'fa-field' }, h('label', { class: 'stack' }, label, control));
    return [
      field('서식 이름', h('input', { value: m.title, oninput: (ev) => (m.title = ev.target.value) })),
      field('제출 기관', h('input', { value: m.agency, list: 'agencyList', oninput: (ev) => (m.agency = ev.target.value.trim() || COMMON_AGENCY) })),
      field('지원 종류', select(KINDS, m.kind, (v) => (m.kind = v))),
      field('구분', select(APPLICANTS, m.applicant, (v) => (m.applicant = v))),
      field('자주 쓰는 서식(★)', h('input', { type: 'checkbox', checked: m.favorite, onchange: (ev) => (m.favorite = ev.target.checked) })),
    ];
  }

  /** 필드 하나의 연결 줄: 자동(사전에 있을 때) · 고정 글자 · 출력할 때 입력 · 비워 둠 */
  function mappingRow(f, mapping) {
    const known = fieldOf(f.name);
    const cur = mapping[f.name] || (known ? { type: 'auto' } : { type: 'ask' });
    mapping[f.name] = cur;
    const text = h('input', { value: cur.text || '', placeholder: '고정 글자', hidden: cur.type !== 'fixed', 'aria-label': `${f.name} 고정 글자`, oninput: (ev) => (cur.text = ev.target.value) });
    const options = known ? [['auto', `자동: ${sourceLabel(known)}에서`], ...MAPPING_TYPES] : MAPPING_TYPES;
    const sel = select(
      options,
      cur.type,
      (v) => {
        cur.type = v;
        text.hidden = v !== 'fixed';
      },
      { 'aria-label': `${f.name} 채울 값` }
    );
    return h('tr', {}, h('td', {}, `{{${f.name}}}`, known ? null : h('div', { class: 'sub', style: 'color: var(--fa-warn)' }, '사전에 없는 필드')), h('td', {}, `${f.count}곳`), h('td', {}, sel, ' ', text));
  }

  function openReview(id) {
    const e = app.library.get(id);
    reviewing = { id, mapping: JSON.parse(JSON.stringify(e.mapping || {})), meta: { title: e.title, agency: e.agency, kind: e.kind, applicant: e.applicant, favorite: e.favorite } };
    replace($('reviewMeta'), metaFields(reviewing.meta));
    fillAgencyList(app.library.agencies().map((a) => a.name));
    $('reviewTitle').textContent = `서식 검토 — ${e.sourceName}`;
    replace($('reviewDocs'), e.documents.length ? e.documents.map(docLine).join(' · ') : '(목차를 찾지 못했습니다)');
    replace(
      $('reviewFields'),
      e.fields.length === 0
        ? h('tr', {}, h('td', { colspan: '3' }, '메일머지 필드가 없는 서식입니다. 출력할 때 자동 인식(머리글 옆 빈 칸 등)으로 채웁니다. [시험 저장]으로 결과를 확인하세요.'))
        : e.fields.map((f) => mappingRow(f, reviewing.mapping))
    );
    $('reviewModal').hidden = false;
  }

  function closeReview() {
    $('reviewModal').hidden = true;
    reviewing = null;
  }

  $('reviewCancel').addEventListener('click', closeReview);
  $('reviewSave').addEventListener('click', () => {
    if (!reviewing) return;
    app.library.update(reviewing.id, { ...reviewing.meta, mapping: reviewing.mapping, reviewed: true });
    log(`서식 검토를 저장했습니다: ${reviewing.meta.title}`, 'ok');
    closeReview();
    changed();
  });

  /* ---------- 시험 저장 · 삭제 ---------- */

  async function testSave(id) {
    const e = app.library.get(id);
    const out = await ipcRenderer.invoke('dialog:saveFile', { title: '가짜 예시 값으로 채운 시험 파일 저장', defaultPath: `${e.title}_시험.hwpx`, filters: [{ name: '한글(HWPX) 문서', extensions: ['hwpx'] }] });
    if (!out) return;
    try {
      const asks = Object.fromEntries(e.fields.map((f) => [f.name, `[${f.name}]`]));
      const r = await saveFilled(app.library, e, SAMPLE_INPUT, out, { asks, hancom });
      log(`시험 파일을 저장했습니다(${r.mode === 'fields' ? `필드 ${r.filled}곳` : `자동 인식 ${r.filled}곳`}): ${out}`, 'ok');
      alert(`가짜 예시 값으로 채운 시험 파일을 저장했습니다.\n한글에서 열어 글자 위치를 확인하세요.\n${out}`);
    } catch (err) {
      notifyError('시험 파일을 만들지 못했습니다', err);
    }
    render(); // 처음 시험 저장하며 한글로 바꿨으면 상태가 바뀐다
  }

  function removeEntry(id) {
    const e = app.library.get(id);
    if (!confirm(`"${e.title}"(${e.agency} · ${e.applicant}) 서식을 목록에서 지울까요? 원본 파일은 그대로 있습니다.`)) return;
    app.library.remove(id);
    log(`서식을 지웠습니다: ${e.title}`);
    changed();
  }

  /* ---------- 백업 ---------- */

  $('btnExport').addEventListener('click', async () => {
    const out = await ipcRenderer.invoke('dialog:saveFile', { title: '서식 백업 내보내기', defaultPath: `서식백업_${new Date().toISOString().slice(0, 10)}.zip`, filters: BACKUP_FILTERS });
    if (!out) return;
    try {
      fs.writeFileSync(out, app.library.exportBackup());
      log(`서식 ${app.library.list().length}개를 백업했습니다: ${out}`, 'ok');
    } catch (err) {
      notifyError('백업을 만들지 못했습니다', err);
    }
  });
  $('btnImport').addEventListener('click', async () => {
    const paths = await ipcRenderer.invoke('dialog:openFile', { title: '서식 백업 가져오기', filters: BACKUP_FILTERS });
    if (!paths || !paths.length) return;
    try {
      const r = app.library.importBackup(fs.readFileSync(paths[0]));
      log(`백업을 가져왔습니다: 새 서식 ${r.added}개, 바뀐 서식 ${r.replaced}개.`, 'ok');
      changed();
    } catch (err) {
      notifyError('백업을 가져오지 못했습니다', err);
    }
  });

  app.on('library', render);
  render();
}

module.exports = { init };
