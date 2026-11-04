// renderer.js
// 격리된 preload 세계의 앱 진입점(preload.js가 index.html에서만 불러온다): 공통 모듈(common.js: 진행 기록·탭 전환) →
// 탭1(문서 양식 변환기) → 탭2(서식 자동 기입: 서류 출력 · 서식 관리 · 설정) 순서로 불러와 각 모듈이 자신의 DOM/이벤트를
// 직접 바인딩하도록 한다. 탭들은 서로를 require하지 않고 common.js에만 의존하므로 common.js를 가장 먼저 불러온다.

const common = require('./common.js');
require('./tab1.js');

common.log('프로그램이 준비되었습니다.', 'ok');
require('./app-shell.js')
  .startApp()
  .catch((err) => common.notifyError('서식 자동 기입을 시작하지 못했습니다', err));
