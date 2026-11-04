// app-shell.js
// 서식 자동 기입 앱의 틀: 화면 전환(서류 출력 · 서식 관리 · 설정), 이 PC의 서식 보관함과 설정 열기,
// 시작할 때 남은 임시 파일 정리. 화면마다의 동작은 print-screen.js · forms-screen.js · settings-screen.js.

const { ipcRenderer } = require('electron');
const path = require('path');
const { openLibrary } = require('../autofill/form-library.js');
const { loadSettings } = require('../autofill/settings.js');
const { cleanupTemp } = require('../autofill/print-job.js');
const { runHancom } = require('../autofill/hancom.js');
const { log } = require('./common.js');

/** 화면들이 함께 쓰는 상태: 보관함·설정, 한글 실행기, 그리고 바뀌었을 때 알림 */
const app = {
  library: null,
  settings: null,
  userDataPath: null,
  /** 한글 자동 조작. 회귀 시나리오는 app.hooks.hancom을 가짜 인쇄로 바꿔 끼운다 */
  hooks: { hancom: runHancom },
  hancom(action, args) {
    return this.hooks.hancom(action, args);
  },
  listeners: { library: [], settings: [] },
  on(kind, fn) {
    this.listeners[kind].push(fn);
  },
  changed(kind) {
    this.listeners[kind].forEach((fn) => fn());
  },
};

const SCREENS = ['screenPrint', 'screenForms', 'screenSettings'];

function showScreen(id) {
  document.querySelectorAll('.fa-nav button').forEach((b) => {
    if (b.dataset.screen === id) b.setAttribute('aria-current', 'page');
    else b.removeAttribute('aria-current');
  });
  SCREENS.forEach((s) => {
    document.getElementById(s).hidden = s !== id;
  });
}

async function startApp() {
  document.querySelectorAll('.fa-nav button').forEach((b) => b.addEventListener('click', () => showScreen(b.dataset.screen)));
  app.userDataPath = await ipcRenderer.invoke('app:getUserDataPath');
  app.library = openLibrary(path.join(app.userDataPath, 'forms'));
  if (app.library.loadError) log(`서식 목록 파일이 망가져 빈 목록으로 시작합니다(원래 파일은 index.json.broken-…으로 남겨 두었습니다): ${app.library.loadError.message}`, 'warn');
  const loaded = loadSettings(app.userDataPath);
  app.settings = loaded.settings;
  if (loaded.error) log('설정 파일을 읽지 못해 기본값으로 시작합니다.', 'warn');
  try {
    const n = cleanupTemp();
    if (n > 0) log(`지난번에 남은 임시 파일 ${n}개를 지웠습니다.`);
  } catch (err) {
    /* 임시 폴더를 못 열어도 시작은 한다 */
  }
  require('./print-screen.js').init(app);
  require('./forms-screen.js').init(app, { showScreen });
  require('./settings-screen.js').init(app);
  showScreen(app.library.list().length > 0 ? 'screenPrint' : 'screenForms');
  log(`서식 ${app.library.list().length}개를 불러왔습니다.`, 'ok');
}

module.exports = { app, startApp, showScreen };
