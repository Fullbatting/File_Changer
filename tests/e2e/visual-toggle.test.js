// tests/e2e/visual-toggle.test.js
// 시각화 켜기/끄기: 기본값 꺼짐, 끄면 매핑 표만, 켜면 작업대 전체, 설정 기억, 끈 상태의 변환 확인 창 미리보기.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/visual-toggle.test.js

const fs = require('fs');
const path = require('path');
const { createHarness } = require('./harness.js');
const { makeSmall } = require('./fixtures.js');

const h = createHarness('visual-toggle');
const { js, sleep, click, record } = h;

const shown = () =>
  js(`({
    toggle: document.getElementById('wbVisualToggle').checked,
    state: document.getElementById('wbVisualState').textContent,
    list: !document.getElementById('wbListView').hidden,
    connect: !document.getElementById('wbConnect').hidden,
    preview: !document.getElementById('wbPreview').hidden,
    summary: !document.getElementById('wbSummary').hidden,
    note: document.getElementById('wbOffNote').hidden ? '' : document.getElementById('wbOffNote').textContent,
  })`);
const setToggle = async (on) => {
  await js(`const t = document.getElementById('wbVisualToggle'); t.checked = ${on}; t.dispatchEvent(new Event('change')); true;`);
  await sleep(150);
};
const settings = () => JSON.parse(fs.readFileSync(path.join(h.userDataDir, 'ui-settings.json'), 'utf-8'));

async function run() {
  await h.start({ visual: null });
  const fx = makeSmall(h.workDir);

  let s = await shown();
  record('설정 파일이 없으면 기본값은 시각화 꺼짐(OFF)', s.toggle === false && s.state === 'OFF', s);
  await h.loadFiles(fx.source, fx.target);
  s = await shown();
  record('꺼짐: 매핑 표만 보이고 요약·연결 보기·미리보기는 숨김', s.list && !s.connect && !s.preview && !s.summary, s);
  const mapping = await js(`Object.fromEntries(Array.from(document.querySelectorAll('#mappingTableBody select')).map((x) => [x.dataset.bcolumn, x.value]))`);
  record('꺼짐에서도 자동 매핑은 그대로', mapping.성명 === '이름' && mapping.입사릴 === '입사일' && mapping.연락처 === '', mapping);
  record('꺼짐: 3단계 설명은 "매핑에 따라 저장"', (await js(`document.getElementById('convertDesc').textContent`)).startsWith('매핑에 따라 4행을 저장합니다'), await js(`document.getElementById('convertDesc').textContent`));

  h.queues.save.push(path.join(h.workDir, 'off.xlsx'));
  await click('btnConvert', 300);
  const modal = await js(`({ preview: !document.getElementById('convertReviewPreview').hidden, head: Array.from(document.querySelectorAll('#convertReviewTableHead th')).map((t) => t.textContent), first: Array.from(document.querySelectorAll('#convertReviewTableBody tr:first-child td')).map((t) => t.textContent) })`);
  record('꺼짐: 변환 확인 창에 저장될 표 앞부분 미리보기', modal.preview && modal.head.join() === '성명,사번,입사릴,급여,부서명,연락처' && modal.first.join() === '홍길동,20230001,2026-01-15,1,500,000,영업팀,', modal);
  await click('convertReviewCancel', 100);

  await setToggle(true);
  s = await shown();
  record('켜기: 요약·연결 보기·미리보기 표시, ON 표시', s.toggle && s.state === 'ON' && s.connect && s.preview && s.summary && !s.list, s);
  record('켜기/끄기 상태를 설정 파일에 저장', settings().visual === true, settings());
  await click('btnConvert', 300);
  record('켜짐: 변환 확인 창은 요약만(미리보기는 화면에 있음)', await js(`document.getElementById('convertReviewPreview').hidden`), '');
  await click('convertReviewCancel', 100);

  await js(`const td = document.querySelector('td.wb-cell[data-r="0"][data-b="성명"]'); td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); const i = document.getElementById('wbCellInput'); i.value = '홍길동(본사)'; i.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); true;`);
  await sleep(150);
  await setToggle(false);
  s = await shown();
  record('끄기: 다시 매핑 표만, 고친 값이 저장에 포함된다는 안내', s.list && !s.connect && !s.preview && s.note.includes('직접 고친 값 1개'), s);
  await js(`document.body.focus(); document.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', code: 'KeyZ', ctrlKey: true, bubbles: true })); true;`);
  await sleep(150);
  record('꺼진 상태에서도 Ctrl+Z 실행 취소 동작(안내 사라짐)', (await shown()).note === '', await shown());
  record('끈 상태도 저장', settings().visual === false, settings());

  await setToggle(true);
  h.win.webContents.reload();
  await sleep(1200);
  await js(`window.alert = () => {}; true;`);
  s = await shown();
  record('다시 시작해도 마지막 켜기/끄기 상태(ON)를 기억', s.toggle === true && s.state === 'ON', s);
}

run()
  .then(() => h.finish())
  .catch((err) => {
    console.log(`ERROR ${err && err.stack ? err.stack : err}`);
    h.finish();
  });
