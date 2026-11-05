// renderer.js
// 격리된 preload 세계의 앱 진입점(preload.js가 index.html에서만 불러온다): 진행 기록(common.js) →
// 사회사업실 도우미(app-shell.js: 서류 출력 · 서식 관리 · 설정).

const common = require('./common.js');
const { startApp } = require('./app-shell.js');

// 화면 코드에서 처리하지 못한 오류도 오류 창으로(검은 창이 없으니 그냥 묻히지 않게). 같은 내용은 잠깐 동안 한 번만
const recent = new Map();
function reportUnexpected(err) {
  const key = String((err && err.message) || err);
  const now = Date.now();
  if (recent.has(key) && now - recent.get(key) < 5000) return;
  recent.set(key, now);
  common.notifyError('예상하지 못한 오류가 발생했습니다', err);
}
window.addEventListener('error', (e) => reportUnexpected(e.error || e.message));
window.addEventListener('unhandledrejection', (e) => reportUnexpected(e.reason));

common.log('프로그램이 준비되었습니다.', 'ok');
startApp().catch((err) => common.notifyError('프로그램을 시작하지 못했습니다', err));
