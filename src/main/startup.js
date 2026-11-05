// src/main/startup.js
// 시작 알림과 오류 창. 프로그램은 run.bat의 검은 창(cmd) 없이 돌기 때문에:
// - 화면이 제대로 떴으면 시작 프로그램(src/launcher/launch.js)에 "떴다"고 알린다 → 그때 검은 창이 닫힌다.
// - 문제가 생기면 오류 창으로 알리고, 아직 시작 중이었으면 시작 프로그램에도 실패를 알린다.
// 알림은 시작 프로그램이 환경 변수(SWH_READY_FILE)로 넘겨 준 파일에 JSON 한 줄로 쓴다(직접 실행하면 알림 없음).

const fs = require('fs');
const { app, dialog } = require('electron');

const TITLE = '사회사업실 도우미';
const ERROR_TITLE = `${TITLE} - 오류`;
const UNEXPECTED = '예상하지 못한 오류가 발생했습니다. 같은 일이 되풀이되면 프로그램을 다시 실행하세요.';

let reported = false;

/** 시작 결과를 한 번만 알린다. result: { ok, error?, shown? } */
function reportLaunch(result) {
  const file = process.env.SWH_READY_FILE;
  if (!file || reported) return;
  reported = true;
  try {
    fs.writeFileSync(file, JSON.stringify(result));
  } catch (e) {
    /* 시작 프로그램이 이미 끝났으면 그만 */
  }
}

const recent = new Map();

/** 오류 창. 같은 내용은 5초 안에 한 번만. 아직 시작 중이면 시작 프로그램에 실패도 알린다 */
function showError(message, err) {
  const detail = err ? String((err && err.message) || err) : '';
  const key = `${message}|${detail}`;
  const now = Date.now();
  if (recent.has(key) && now - recent.get(key) < 5000) return;
  recent.set(key, now);
  const text = detail ? `${message}\n\n${detail}` : message;
  reportLaunch({ ok: false, error: text, shown: true });
  try {
    if (app.isReady()) dialog.showErrorBox(ERROR_TITLE, text);
    else console.error(text);
  } catch (e) {
    console.error(text);
  }
}

/** 메인 프로세스에서 처리하지 못한 오류 */
function watchProcess() {
  process.on('uncaughtException', (err) => showError(UNEXPECTED, err));
  process.on('unhandledRejection', (err) => showError(UNEXPECTED, err));
}

/** 창의 문제(화면 준비 실패·불러오기 실패·꺼짐·응답 없음) */
function watchWindow(win) {
  const wc = win.webContents;
  wc.on('preload-error', (e, preloadPath, err) => showError('프로그램 화면을 준비하지 못했습니다. run.bat을 다시 실행해 보세요.', err));
  wc.on('did-fail-load', (e, code, desc, url, isMainFrame) => {
    if (isMainFrame && code !== -3) showError('프로그램 화면을 불러오지 못했습니다.', `${desc} (${code})`);
  });
  wc.on('render-process-gone', async (e, details) => {
    if (details.reason === 'clean-exit' || win.isDestroyed()) return;
    reportLaunch({ ok: false, error: `프로그램 화면이 꺼졌습니다(${details.reason}).`, shown: true });
    const r = await dialog.showMessageBox(win, {
      type: 'error',
      title: ERROR_TITLE,
      message: '프로그램 화면이 갑자기 꺼졌습니다.',
      detail: `입력하던 인적사항은 지워졌습니다. 다시 불러와 처음부터 입력하세요. (${details.reason}, 코드 ${details.exitCode})`,
      buttons: ['다시 불러오기', '프로그램 닫기'],
      defaultId: 0,
      noLink: true,
    });
    if (win.isDestroyed()) return;
    if (r.response === 0) wc.reload();
    else win.close();
  });
  win.on('unresponsive', async () => {
    const r = await dialog.showMessageBox(win, {
      type: 'warning',
      title: TITLE,
      message: '프로그램이 응답하지 않습니다.',
      detail: '한글로 큰 서식을 다루는 중이면 잠시 기다리면 돌아옵니다.',
      buttons: ['기다리기', '프로그램 닫기'],
      defaultId: 0,
      noLink: true,
    });
    if (r.response === 1 && !win.isDestroyed()) win.destroy();
  });
}

module.exports = { reportLaunch, showError, watchProcess, watchWindow, ERROR_TITLE };
