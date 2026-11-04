// renderer.js
// 격리된 preload 세계의 앱 진입점(preload.js가 index.html에서만 불러온다): 진행 기록(common.js) →
// 서식 자동 기입 앱(app-shell.js: 서류 출력 · 서식 관리 · 설정).

const common = require('./common.js');
const { startApp } = require('./app-shell.js');

common.log('프로그램이 준비되었습니다.', 'ok');
startApp().catch((err) => common.notifyError('프로그램을 시작하지 못했습니다', err));
