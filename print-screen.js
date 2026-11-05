// print-screen.js
// ① 서류 출력 화면: 인적사항(고른 서류에 있는 칸만) → 제출 기관(주소로 추천) → 세대주/세대원 → 서류 고르기 → 출력.
// 인적사항은 화면(메모리)에만 있고 저장하지 않는다. 출력이 끝나면 지운다. 로그에는 값을 남기지 않는다.
// 무엇을 보이고 무엇이 비었는지의 규칙은 print-plan.js에 있다.

const { ipcRenderer } = require('electron');
const fs = require('fs');
const path = require('path');
const { h, replace } = require('./ui-dom.js');
const { INPUT_BY_KEY, INPUT_GROUPS, checkRrn, rrnSummary } = require('./merge-fields.js');
const { suggestAgencies, entryStatus } = require('./form-library.js');
const { printForms, saveFilled } = require('./print-job.js');
const plan = require('./print-plan.js');
const { log } = require('./common.js');

const RRN_HINT = '생년월일(6자리)·{{1}}~{{13}}은 여기서 자동으로 만듭니다';

function init(app) {
  const $ = (id) => document.getElementById(id);
  const state = {
    input: {}, // 인적사항(저장 안 함)
    asks: {}, // 사전에 없는 필드 값(저장 안 함)
    agency: null,
    agencyPicked: false, // 사용자가 기관을 직접 골랐으면 주소가 바뀌어도 바꾸지 않는다
    suggested: [], // 주소에 맞는 기관들
    applicant: '세대주',
    checked: new Set(),
    copies: app.settings.copies || 1,
  };
  const inputEls = new Map();
  const fieldBox = (key) => inputEls.get(key).closest('.fa-field');

  const checkedEntries = () => app.library.list().filter((e) => state.checked.has(e.id));
  // 서식에 들어갈 입력: 화면의 인적사항 + 설정의 담당자
  const fullInput = () => ({ ...state.input, ...Object.fromEntries(INPUT_GROUPS.staff.map((k) => [k, app.settings[k]])) });
  const pickNothing = () => state.checked.clear();

  /* ---------- 1. 인적사항 칸 ---------- */

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
    return h(
      'div',
      { class: `fa-field${key === 'addr' ? ' wide' : ''}`, 'data-key': key },
      h('label', { for: id }, def.label, def.required ? h('span', { class: 'fa-req' }, '필수') : null),
      input,
      h('span', { class: 'fa-hint', id: `${id}-hint` }, def.hint || '')
    );
  }

  function showRrn() {
    const v = state.input.rrn || '';
    const problem = checkRrn(v);
    inputEls.get('rrn').setAttribute('aria-invalid', problem ? 'true' : 'false');
    const hint = $('in-rrn-hint');
    hint.className = problem ? 'fa-hint bad' : 'fa-hint';
    hint.textContent = problem || rrnSummary(v) || RRN_HINT;
  }

  /** 고른 서류에 따라: 진료 정보 칸, 쓰지 않는 칸 흐리게, 출력 때 입력 칸, 설정의 담당자 안내 */
  function renderNeeds() {
    const list = checkedEntries();
    const keys = plan.neededInputs(list);
    const care = INPUT_GROUPS.care.filter((k) => keys.includes(k));
    INPUT_GROUPS.care.forEach((k) => {
      fieldBox(k).hidden = !care.includes(k);
    });
    $('careBlock').hidden = care.length === 0;
    INPUT_GROUPS.patient.forEach((k) => fieldBox(k).classList.toggle('unused', list.length > 0 && !keys.includes(k)));

    const asks = plan.askFieldNames(list);
    $('askBlock').hidden = asks.length === 0;
    replace(
      $('askFields'),
      asks.map((name, i) => {
        const input = h('input', {
          id: `ask-${i}`,
          value: state.asks[name] || '',
          autocomplete: 'off',
          oninput: () => {
            state.asks[name] = input.value;
            renderFormList();
            renderSummary();
          },
        });
        return h('div', { class: 'fa-field' }, h('label', { for: `ask-${i}` }, `{{${name}}}`), input);
      })
    );

    const staff = plan.missingStaff(list, app.settings);
    $('staffNote').hidden = staff.length === 0;
    replace(
      $('staffNote'),
      `고른 서류에 ${staff.join('·')} 칸이 있습니다. `,
      h('button', { type: 'button', class: 'fa-link', onclick: () => document.querySelector('[data-screen="screenSettings"]').click() }, '설정에서 넣기')
    );
  }

  /* ---------- 2. 제출 기관 · 3. 신청 구분 ---------- */

  function suggest() {
    state.suggested = suggestAgencies(state.input.addr, app.library.agencies().map((a) => a.name));
    if (!state.agencyPicked && state.suggested.length > 0 && state.agency !== state.suggested[0]) {
      state.agency = state.suggested[0];
      pickNothing();
    }
  }

  function pickAgency(name) {
    state.agency = name;
    state.agencyPicked = true;
    pickNothing();
    refresh();
  }

  function renderAgencies() {
    const agencies = app.library.agencies();
    if (!agencies.some((a) => a.name === state.agency)) state.agency = agencies.length ? agencies[0].name : null;
    replace(
      $('agencyChips'),
      agencies.length === 0
        ? h('div', { class: 'fa-empty' }, '등록한 서식이 없습니다. [서식 관리]에서 서식을 먼저 등록하세요.')
        : agencies.map((a) =>
            h(
              'button',
              { type: 'button', class: 'fa-chip', 'aria-pressed': String(a.name === state.agency), onclick: () => pickAgency(a.name) },
              a.name,
              h('span', { class: 'count' }, String(a.count)),
              state.suggested.includes(a.name) ? h('span', { class: 'tag' }, '주소로 추천') : null
            )
          )
    );
    $('agencyNote').textContent = state.suggested.length > 0 ? `주소에 맞는 기관: ${state.suggested.join(', ')}` : state.input.addr ? '주소에 맞는 기관이 없습니다. 직접 고르세요' : '';
  }

  document.querySelectorAll('input[name="applicant"]').forEach((r) =>
    r.addEventListener('change', () => {
      state.applicant = r.value;
      pickNothing();
      refresh();
    })
  );

  /* ---------- 4. 출력할 서류 ---------- */

  function formBadge(entry, on) {
    if (!entryStatus(entry).ready) return h('span', { class: 'fa-badge warn' }, '서식 관리에서 확인 필요');
    if (!on) return h('span', { class: 'fa-badge muted' }, `서류 ${entry.documents.length || '?'}종`);
    const miss = plan.missingFor(entry, fullInput(), state.asks);
    return miss.length > 0 ? h('span', { class: 'fa-badge warn' }, `빈칸: ${miss.join(', ')}`) : h('span', { class: 'fa-badge' }, '모두 채워짐');
  }

  function formItem(entry) {
    const id = `pick-${entry.id}`;
    const on = state.checked.has(entry.id);
    const docs = entry.documents.map((d) => d.title || d.mark).filter(Boolean).join(', ');
    return h(
      'li',
      { class: `fa-form${on ? ' checked' : ''}` },
      h('input', {
        type: 'checkbox',
        id,
        checked: on,
        disabled: !entryStatus(entry).ready,
        onchange: (ev) => {
          if (ev.target.checked) state.checked.add(entry.id);
          else state.checked.delete(entry.id);
          refresh();
        },
      }),
      h(
        'label',
        { for: id },
        h('div', { class: 'name' }, `${entry.favorite ? '★ ' : ''}${entry.title}`, entry.applicant !== '공통' ? ` (${entry.applicant})` : ''),
        h('div', { class: 'docs' }, docs || (entry.fields.length ? '' : '메일머지 필드가 없어 자동 인식으로 채웁니다'))
      ),
      formBadge(entry, on)
    );
  }

  function renderFormList() {
    const list = plan.entriesFor(app.library.list(), state.agency, state.applicant);
    // 하나뿐이면 미리 골라 두고, 목록에 없는 것은 고른 것에서 뺀다
    if (list.length === 1 && state.checked.size === 0 && entryStatus(list[0]).ready) state.checked.add(list[0].id);
    [...state.checked].forEach((id) => {
      if (!list.some((e) => e.id === id)) state.checked.delete(id);
    });
    $('formsNote').textContent = state.agency ? `${state.agency} ${state.applicant}용 서식 ${list.length}개` : '';
    if (list.length === 0) {
      replace($('formList'), h('div', { class: 'fa-empty' }, state.agency ? `${state.agency}에 ${state.applicant}용 서식이 없습니다.` : '제출 기관을 고르세요.'));
      return;
    }
    replace(
      $('formList'),
      plan.groupByKind(list).map((g) => [h('p', { class: 'fa-kind' }, g.kind), h('ul', { class: 'fa-forms' }, g.items.map(formItem))])
    );
  }

  /* ---------- 출력 줄 ---------- */

  function renderSummary() {
    const list = checkedEntries();
    const s = plan.summaryOf(list, fullInput(), state.asks, state.copies);
    replace($('printSummary'), h('strong', {}, s.title), s.detail);
    $('copiesValue').textContent = String(state.copies);
    $('btnPrint').querySelector('span').textContent = list.length > 0 ? `서식 출력 (${list.length}건)` : '서식 출력';
    $('btnPrint').disabled = list.length === 0;
    $('btnSaveFilled').disabled = list.length === 0;
  }

  function refresh() {
    renderAgencies();
    renderFormList();
    renderNeeds();
    renderSummary();
  }

  const setCopies = (n) => {
    state.copies = Math.max(1, Math.min(99, n));
    renderSummary();
  };
  $('btnFewer').addEventListener('click', () => setCopies(state.copies - 1));
  $('btnMore').addEventListener('click', () => setCopies(state.copies + 1));

  /** 인적사항을 모두 지운다(화면 칸·메모리) */
  function clearPerson() {
    state.input = {};
    state.asks = {};
    inputEls.forEach((el) => {
      el.value = '';
    });
    state.agencyPicked = false;
    state.suggested = [];
    showRrn();
    refresh();
  }
  $('btnClearPerson').addEventListener('click', clearPerson);

  /** 출력 전 확인(막을 이유가 있으면 알리고 그 칸으로, 빈칸이 있으면 계속할지). 계속하면 true */
  function confirmBeforeOutput(list) {
    const check = plan.checkBeforeOutput(list, fullInput(), state.asks);
    if (check.block) {
      alert(check.block);
      inputEls.get(check.focus).focus();
      return false;
    }
    return check.missing.length === 0 || confirm(`다음 칸은 비운 채로 나갑니다: ${check.missing.join(', ')}\n계속할까요?`);
  }

  /* ---------- 출력 진행 창 ---------- */

  const progress = {
    open(title) {
      $('printTitle').textContent = title;
      $('printProgress').style.width = '0%';
      $('printStatus').textContent = '';
      replace($('printResults'));
      $('printClose').hidden = true;
      $('printModal').hidden = false;
    },
    step(i, total, entry) {
      $('printProgress').style.width = `${Math.round((i / total) * 100)}%`;
      $('printStatus').textContent = `${i + 1}/${total}: ${entry.title}`;
    },
    done(results, note) {
      $('printProgress').style.width = '100%';
      $('printStatus').textContent = note;
      replace($('printResults'), results.map((r) => h('li', { class: r.ok ? '' : 'bad' }, r.ok ? `✓ ${r.title}` : `✗ ${r.title}: ${r.error}`)));
      $('printClose').hidden = false;
      $('printClose').focus();
    },
  };
  $('printClose').addEventListener('click', () => {
    $('printModal').hidden = true;
  });

  /* ---------- 출력 · 한글 파일로 저장 ---------- */

  $('btnPrint').addEventListener('click', async () => {
    const list = checkedEntries();
    if (list.length === 0 || !confirmBeforeOutput(list)) return;
    progress.open('출력 중');
    const results = await printForms(app.library, list, fullInput(), {
      copies: state.copies,
      printer: app.settings.printer,
      asks: state.asks,
      hancom: (action, args) => app.hancom(action, args),
      onProgress: (i, total, e) => progress.step(i, total, e),
    });
    const failed = results.filter((r) => !r.ok).length;
    log(`서식 ${results.length}건 출력: ${results.length - failed}건 성공${failed ? `, ${failed}건 실패` : ''}.`, failed ? 'warn' : 'ok');
    if (failed === 0) {
      clearPerson();
      progress.done(results, '출력을 마쳤습니다. 인적사항을 지웠습니다.');
    } else {
      progress.done(results, '일부 서류를 출력하지 못했습니다. 인적사항은 다시 출력할 수 있게 남겨 두었습니다. 다 끝나면 [새로 입력]으로 지우세요.');
    }
    app.changed('library'); // 처음 출력하며 한글로 바꾼 서식이 있으면 상태가 바뀐다
  });

  /** 폴더 안에서 겹치지 않는 파일 경로("이름(2).hwpx" …) */
  const freePath = (dir, name) => {
    let out = path.join(dir, name);
    for (let n = 2; fs.existsSync(out); n++) out = path.join(dir, name.replace(/\.hwpx$/, `(${n}).hwpx`));
    return out;
  };

  $('btnSaveFilled').addEventListener('click', async () => {
    const list = checkedEntries();
    if (list.length === 0 || !confirmBeforeOutput(list)) return;
    if (!confirm('채운 서식을 한글 파일(.hwpx)로 저장합니다. 저장한 파일에는 인적사항이 들어 있으니, 다 쓴 뒤 지워 주세요.\n저장할 폴더를 고를까요?')) return;
    const dir = await ipcRenderer.invoke('dialog:selectFolder', { title: '채운 서식을 저장할 폴더' });
    if (!dir) return;
    progress.open('한글 파일로 저장 중');
    const results = [];
    for (let i = 0; i < list.length; i++) {
      const e = list[i];
      progress.step(i, list.length, e);
      const out = freePath(dir, plan.filledFileName(e));
      try {
        await saveFilled(app.library, e, fullInput(), out, { asks: state.asks, hancom: (a, b) => app.hancom(a, b) });
        results.push({ title: `${e.title} → ${path.basename(out)}`, ok: true });
      } catch (err) {
        results.push({ title: e.title, ok: false, error: err.message });
      }
    }
    log(`채운 서식 ${results.filter((r) => r.ok).length}건을 저장했습니다: ${dir}`, 'ok');
    progress.done(results, `저장 위치: ${dir}`);
  });

  /* ---------- 처음 · 바뀔 때 ---------- */

  function renderPrinterNote() {
    $('printerNote').textContent =
      process.platform === 'win32'
        ? `프린터: ${app.settings.printer || '기본 프린터'}. 한글로 바로 인쇄합니다(채운 임시 파일은 인쇄 뒤 바로 지웁니다)`
        : '이 PC에서는 한글로 바로 인쇄할 수 없습니다(Windows와 한글이 필요). [한글 파일로 저장]을 쓰세요.';
  }

  replace($('personFields'), INPUT_GROUPS.patient.map(fieldEl));
  replace($('careFields'), INPUT_GROUPS.care.map(fieldEl));
  showRrn();
  app.on('library', () => {
    suggest();
    refresh();
  });
  app.on('settings', () => {
    state.copies = app.settings.copies || state.copies;
    renderPrinterNote();
    refresh();
  });
  renderPrinterNote();
  refresh();
}

module.exports = { init };
