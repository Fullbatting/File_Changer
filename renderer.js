// renderer.js
// 렌더러 프로세스 진입점: 화면 공통 도우미(common.js: 로그·확인 창·동의어 사전 창) → 예전 메일머지 화면(tab2.js) →
// 서식 자동 기입 앱(app-shell.js: 서류 출력 · 서식 관리 · 설정) 순서로 불러온다.

const common = require('./common.js');
require('./tab2.js');
const { startApp } = require('./app-shell.js');

common.log('프로그램이 준비되었습니다.', 'ok');
startApp().catch((err) => common.notifyError('프로그램을 시작하지 못했습니다', err));
