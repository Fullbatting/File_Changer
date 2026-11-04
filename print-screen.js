// print-screen.js
// ① 서류 출력 화면: 인적사항(고른 서류에 있는 칸만) → 제출 기관(주소로 추천) → 세대주/세대원 → 서류 고르기 → 출력.
// 인적사항은 화면(메모리)에만 있고 저장하지 않는다. 출력이 끝나면 지운다. 로그에는 값을 남기지 않는다.

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const { h, replace } = require('./ui-dom.js');
const { INPUT_BY_KEY, inputsFor, missingInputs, checkRrn, rrnSummary } = require('./merge-fields.js');
const { KINDS, suggestAgencies, entryStatus } = require('./form-library.js');
const { printForms, saveFilled } = require('./print-job.js');
const { runHancom } = require('./hancom.js');
const { log } = require('./common.js');

/** 바꿔 끼울 수 있는 한글 실행기(회귀 시나리오가 가짜 인쇄로 바꾼다) */
const hooks = { hancom: runHancom };

const PERSON_KEYS = ['name', 'rrn', 'addr', 'phone', 'writtenDate'];
const CARE_KEYS = ['careStart', 'careEnd', 'disease', 'diseaseCode', 'diagnosisDate'];
const STAFF_KEYS = ['staffName', 'staffTitle', 'staffPhone'];

function init(app) {
  const $ = (id) => document.getElementById(id);
  const state = {
    input: {}, // 인적사항(저장 안 함)
    asks: {}, // 사전에 없는 필드 값(저장 안 함)
    agency: null,
    agencyPicked: false, // 사용자가 기관을 직접 골랐으면 주소가 바뀌어도 바꾸지 않는다
    applicant: '세대주',
    checked: new Set(),
    copies: app.settings.copies || 1,
  };
  const inputEls = new Map();

  /* ---------- 인적사항 칸 ---------- */

  function fieldEl(key) {
    const def = INPUT_BY_KEY.get(key);
    const id = `in-${key}`;
    const input = h('input', {
      id,
      type: def.type === 'date' ? 'date' : 'text',
      inputmode: def.type === 'tel' || def.type === 'rrn' ? 'numeric' : null,
      placeholder: def.placeholder || null,
      autocomplete: 'off',
      spellcheck: 'false',
      oninput: () => {
        state.input[key] = input.value;
        if (key === 'rrn') showRrn();
        if (key === 'addr') suggest();
        refresh();
      },
    });
    inputEls.set(key, input);
    const hint = h('span', { class: 'fa-hint', id: `${id}-hint` }, def.hint || '');
    return h('div', { class: `fa-field${key === 'addr' ? ' wide' : ''}`, 'data-key': key }, h('label', { for: id }, def.label, def.required ? h('span', { class: 'fa-req' }, '필수') : null), input, hint);
  }

  function showRrn() {
    const v = state.input.rrn || '';
    const problem = checkRrn(v);
    const el = inputEls.get('rrn');
    el.setAttribute('aria-invalid', problem ? 'true' : 'false');
    const hint = $('in-rrn-hint');
    hint.className = problem ? 'fa-hint bad' : 'fa-hint';
    hint.textContent = problem || rrnSummary(v) || '생년월일(6자리)·{{1}}~{{13}}은 여기서 자동으로 만듭니다';
  }

  replace($('personFields'), PERSON_KEYS.map(fieldEl));
  replace($('careFields'), CARE_KEYS.map(fieldEl));
  showRrn();

  /* ---------- 고른 서류에 필요한 것 ---------- */

  const entries = () => app.library.list();
  const checkedEntries = () => entries().filter((e) => state.checked.has(e.id));
  const autoNames = (list) => list.flatMap((e) => e.fields.filter((f) => !e.mapping[f.name] || e.mapping[f.name].type === 'auto').map((f) => f.name));
  const askNames = (list) => [...new Set(list.flatMap((e) => e.fields.filter((f) => (e.mapping[f.name] || {}).type === 'ask').map((f) => f.name)))];
  const fullInput = () => ({ ...state.input, staffName: app.settings.staffName, staffTitle: app.settings.staffTitle, staffPhone: app.settings.staffPhone });

  /** 서식 하나에 비어 있는 입력(이름표) */
  function missingFor(entry) {
    const { keys } = inputsFor(autoNames([entry]));
    const miss = missingInputs(keys, fullInput()).map((k) => INPUT_BY_KEY.get(k).label);
    askNames([entry]).forEach((n) => {
      if (!String(state.asks[n] || '').trim()) miss.push(n);
    });
    return miss;
  }

  function renderNeeds() {
    const list = checkedEntries();
    const { keys } = inputsFor(autoNames(list));
    // 진료 정보: 고른 서류에 있는 칸만(아무것도 안 골랐으면 숨김)
    let anyCare = false;
    CARE_KEYS.forEach((k) => {
      const show = keys.includes(k);
      anyCare = anyCare || show;
      inputEls.get(k).closest('.fa-field').hidden = !show;
    });
    $('careBlock').hidden = !anyCare;
    // 연락처·주소처럼 서식에 없는 칸은 흐리게만(입력은 해도 됨)
    PERSON_KEYS.forEach((k) => inputEls.get(k).closest('.fa-field').classList.toggle('unused', list.length > 0 && !keys.includes(k)));
    // 사전에 없는 필드
    const asks = askNames(list);
    $('askBlock').hidden = asks.length === 0;
    replace(
      $('askFields'),
      asks.map((name, i) => {
        const id = `ask-${i}`;
        const input = h('input', { id, value: state.asks[name] || '', autocomplete: 'off', oninput: () => ((state.asks[name] = input.value), renderFormList(), renderSummary()) });
        return h('div', { class: 'fa-field' }, h('label', { for: id }, `{{${name}}}`), input);
      })
    );
    // 담당자 정보는 설정에서
    const staffMissing = keys.filter((k) => STAFF_KEYS.includes(k) && !String(app.settings[k] || '').trim());
    const note = $('staffNote');
    note.hidden = staffMissing.length === 0;
    replace(note, `고른 서류에 ${staffMissing.map((k) => INPUT_BY_KEY.get(k).label).join('·')} 칸이 있습니다. `, h('button', { type: 'button', class: 'fa-link', onclick: () => document.querySelector('[data-screen="screenSettings"]').click() }, '설정에서 넣기'));
  }

  /* ---------- 제출 기관 ---------- */

  let suggested = [];
  function suggest() {
    suggested = suggestAgencies(state.input.addr, app.library.agencies().map((a) => a.name));
    if (!state.agencyPicked && suggested.length > 0 && state.agency !== suggested[0]) {
      state.agency = suggested[0];
      state.checked.clear();
    }
    renderAgencies();
  }

  function renderAgencies() {
    const agencies = app.library.agencies();
    if (!state.agency || !agencies.some((a) => a.name === state.agency)) state.agency = agencies.length ? agencies[0].name : null;
    replace(
      $('agencyChips'),
      agencies.length === 0
        ? h('div', { class: 'fa-empty' }, '등록한 서식이 없습니다. [서식 관리]에서 서식을 먼저 등록하세요.')
        : agencies.map((a) =>
            h(
              'button',
              {
                type: 'button',
                class: 'fa-chip',
                'aria-pressed': String(a.name === state.agency),
                onclick: () => {
                  state.agency = a.name;
                  state.agencyPicked = true;
                  state.checked.clear();
                  refresh();
                },
              },
              a.name,
              h('span', { class: 'count' }, String(a.count)),
              suggested.includes(a.name) ? h('span', { class: 'tag' }, '주소로 추천') : null
            )
          )
    );
    $('agencyNote').textContent = suggested.length > 0 ? `주소에 맞는 기관: ${suggested.join(', ')}` : state.input.addr ? '주소에 맞는 기관이 없습니다 — 직접 고르세요' : '';
  }

  document.querySelectorAll('input[name="applicant"]').forEach((r) =>
    r.addEventListener('change', () => {
      state.applicant = r.value;
      state.checked.clear();
      refresh();
    })
  );

  /* ---------- 서류 목록 ---------- */

  function visibleEntries() {
    return entries()
      .filter((e) => e.agency === state.agency && (e.applicant === state.applicant || e.applicant === '공통'))
      .sort((a, b) => KINDS.indexOf(a.kind) - KINDS.indexOf(b.kind) || b.favorite - a.favorite || a.title.localeCompare(b.title, 'ko'));
  }

  function renderFormList() {
    const list = visibleEntries();
    // 하나뿐이면 미리 골라 둔다
    if (list.length === 1 && state.checked.size === 0 && entryStatus(list[0]).ready) state.checked.add(list[0].id);
    [...state.checked].forEach((id) => {
      if (!list.some((e) => e.id === id)) state.checked.delete(id);
    });
    $('formsNote').textContent = state.agency ? `${state.agency} · ${state.applicant} 서식 ${list.length}개` : '';
    if (list.length === 0) {
      replace($('formList'), h('div', { class: 'fa-empty' }, state.agency ? `${state.agency}에 ${state.applicant}용 서식이 없습니다.` : '제출 기관을 고르세요.'));
      return;
    }
    const groups = [];
    list.forEach((e) => {
      const g = groups.find((x) => x.kind === e.kind);
      if (g) g.items.push(e);
      else groups.push({ kind: e.kind, items: [e] });
    });
    replace(
      $('formList'),
      groups.map((g) => [
        h('p', { class: 'fa-kind' }, g.kind),
        h(
          'ul',
          { class: 'fa-forms' },
          g.items.map((e) => {
            const status = entryStatus(e);
            const id = `pick-${e.id}`;
            const on = state.checked.has(e.id);
            const miss = on ? missingFor(e) : [];
            const badge = !status.ready
              ? h('span', { class: 'fa-badge warn' }, '서식 관리에서 확인 필요')
              : !on
                ? h('span', { class: 'fa-badge muted' }, `서류 ${e.documents.length || '?'}종`)
                : miss.length > 0
                  ? h('span', { class: 'fa-badge warn' }, `빈칸: ${miss.join(', ')}`)
                  : h('span', { class: 'fa-badge' }, '모두 채워짐');
            const docs = e.documents.map((d) => d.title || d.mark).filter(Boolean).join(' · ');
            return h(
              'li',
              { class: `fa-form${on ? ' checked' : ''}` },
              h('input', {
                type: 'checkbox',
                id,
                checked: on,
                disabled: !status.ready,
                onchange: (ev) => {
                  if (ev.target.checked) state.checked.add(e.id);
                  else state.checked.delete(e.id);
                  refresh();
                },
              }),
              h('label', { for: id }, h('div', { class: 'name' }, `${e.favorite ? '★ ' : ''}${e.title}`, e.applicant !== '공통' ? ` (${e.applicant})` : ''), h('div', { class: 'docs' }, docs || (e.fields.length ? '' : '메일머지 필드가 없어 자동 인식으로 채웁니다'))),
              badge
            );
          })
        ),
      ])
    );
  }

  /* ---------- 출력 줄 ---------- */

  function renderSummary() {
    const list = checkedEntries();
    const docCount = list.reduce((n, e) => n + (e.documents.length || 1), 0);
    const miss = [...new Set(list.flatMap(missingFor))];
    const rrnProblem = checkRrn(state.input.rrn || '');
    replace(
      $('printSummary'),
      h('strong', {}, list.length === 0 ? '출력할 서류를 고르세요' : `서식 ${list.length}개 · 서류 약 ${docCount}종 × ${state.copies}부`),
      rrnProblem ? `주민등록번호: ${rrnProblem}` : miss.length > 0 ? `빈칸으로 나갈 칸: ${miss.join(', ')}` : list.length > 0 ? '모든 칸이 채워집니다' : ''
    );
    $('copiesValue').textContent = String(state.copies);
    $('btnPrint').textContent = list.length > 0 ? `서식 출력 (${list.length}건)` : '서식 출력';
    $('btnPrint').disabled = list.length === 0;
    $('btnSaveFilled').disabled = list.length === 0;
  }

  function refresh() {
    renderAgencies();
    renderFormList();
    renderNeeds();
    renderSummary();
  }

  $('btnFewer').addEventListener('click', () => {
    state.copies = Math.max(1, state.copies - 1);
    renderSummary();
  });
  $('btnMore').addEventListener('click', () => {
    state.copies = Math.min(99, state.copies + 1);
    renderSummary();
  });

  /** 인적사항을 모두 지운다(화면 칸·메모리) */
  function clearPerson() {
    state.input = {};
    state.asks = {};
    inputEls.forEach((el) => {
      el.value = '';
    });
    state.agencyPicked = false;
    suggested = [];
    showRrn();
    refresh();
  }
  $('btnClearPerson').addEventListener('click', clearPerson);

  /** 출력 전에 확인: 필수 칸·주민등록번호 모양, 빈칸으로 나갈 칸. 계속하면 true */
  function confirmBeforeOutput(list) {
    const { keys } = inputsFor(autoNames(list));
    const need = keys.filter((k) => INPUT_BY_KEY.get(k).required && !String(state.input[k] || '').trim());
    if (need.length > 0) {
      alert(`${need.map((k) => INPUT_BY_KEY.get(k).label).join(', ')}을(를) 넣어 주세요.`);
      inputEls.get(need[0]).focus();
      return false;
    }
    const rrnProblem = checkRrn(state.input.rrn || '');
    if (rrnProblem) {
      alert(`주민등록번호를 확인하세요: ${rrnProblem}`);
      inputEls.get('rrn').focus();
      return false;
    }
    const miss = [...new Set(list.flatMap(missingFor))];
    return miss.length === 0 || confirm(`다음 칸은 비운 채로 나갑니다: ${miss.join(', ')}\n계속할까요?`);
  }

  /* ---------- 출력 ---------- */

  const modal = {
    el: $('printModal'),
    open(title) {
      $('printTitle').textContent = title;
      $('printProgress').style.width = '0%';
      $('printStatus').textContent = '';
      replace($('printResults'));
      $('printClose').hidden = true;
      this.el.hidden = false;
    },
    progress(i, total, text) {
      $('printProgress').style.width = `${Math.round((i / total) * 100)}%`;
      $('printStatus').textContent = text;
    },
    done(results, note) {
      $('printProgress').style.width = '100%';
      $('printStatus').textContent = note;
      replace(
        $('printResults'),
        results.map((r) => h('li', { class: r.ok ? '' : 'bad' }, r.ok ? `✓ ${r.title}` : `✗ ${r.title} — ${r.error}`))
      );
      $('printClose').hidden = false;
      $('printClose').focus();
    },
  };
  $('printClose').addEventListener('click', () => {
    modal.el.hidden = true;
  });

  $('btnPrint').addEventListener('click', async () => {
    const list = checkedEntries();
    if (list.length === 0 || !confirmBeforeOutput(list)) return;
    modal.open('출력 중');
    const results = await printForms(app.library, list, fullInput(), {
      copies: state.copies,
      printer: app.settings.printer,
      asks: state.asks,
      hancom: (action, args) => hooks.hancom(action, args),
      onProgress: (i, total, e) => modal.progress(i, total, `${i + 1}/${total} · ${e.title}`),
    });
    const failed = results.filter((r) => !r.ok).length;
    log(`서식 ${results.length}건 출력: ${results.length - failed}건 성공${failed ? `, ${failed}건 실패` : ''}.`, failed ? 'warn' : 'ok');
    if (failed === 0) {
      clearPerson();
      modal.done(results, '출력을 마쳤습니다. 인적사항을 지웠습니다.');
    } else {
      modal.done(results, '일부 서류를 출력하지 못했습니다. 인적사항은 다시 출력할 수 있게 남겨 두었습니다 — 다 끝나면 [새로 입력]으로 지우세요.');
    }
    app.changed('library'); // 처음 출력하며 한글로 바꾼 서식이 있으면 상태가 바뀐다
  });

  $('btnSaveFilled').addEventListener('click', async () => {
    const list = checkedEntries();
    if (list.length === 0 || !confirmBeforeOutput(list)) return;
    if (!confirm('채운 서식을 한글 파일(.hwpx)로 저장합니다. 저장한 파일에는 인적사항이 들어 있으니, 다 쓴 뒤 지워 주세요.\n저장할 폴더를 고를까요?')) return;
    const dir = await ipcRenderer.invoke('dialog:selectFolder', { title: '채운 서식을 저장할 폴더' });
    if (!dir) return;
    modal.open('한글 파일로 저장 중');
    const results = [];
    for (let i = 0; i < list.length; i++) {
      const e = list[i];
      modal.progress(i, list.length, `${i + 1}/${list.length} · ${e.title}`);
      const name = `${e.agency}_${e.title}${e.applicant !== '공통' ? `(${e.applicant})` : ''}.hwpx`.replace(/[\\/:*?"<>|]/g, '_');
      let out = path.join(dir, name);
      for (let n = 2; fs.existsSync(out); n++) out = path.join(dir, name.replace(/\.hwpx$/, `(${n}).hwpx`));
      try {
        await saveFilled(app.library, e, fullInput(), out, { asks: state.asks, hancom: (a, b) => hooks.hancom(a, b) });
        results.push({ title: `${e.title} → ${path.basename(out)}`, ok: true });
      } catch (err) {
        results.push({ title: e.title, ok: false, error: err.message });
      }
    }
    log(`채운 서식 ${results.filter((r) => r.ok).length}건을 저장했습니다: ${dir}`, 'ok');
    modal.done(results, `저장 위치: ${dir}`);
  });

  /* ---------- 처음 · 바뀔 때 ---------- */

  const printerNote = () => {
    $('printerNote').textContent =
      process.platform === 'win32'
        ? `프린터: ${app.settings.printer || '기본 프린터'} · 한글로 바로 인쇄합니다(채운 임시 파일은 인쇄 뒤 바로 지웁니다)`
        : '이 PC에서는 한글로 바로 인쇄할 수 없습니다(Windows와 한글이 필요). [한글 파일로 저장]을 쓰세요.';
  };
  app.on('library', () => {
    suggest();
    refresh();
  });
  app.on('settings', () => {
    state.copies = app.settings.copies || state.copies;
    printerNote();
    refresh();
  });
  printerNote();
  refresh();
}

module.exports = { init, hooks };
