// renderer.js
// 렌더러 프로세스 진입점: 화면 공통 도우미(common.js: 로그·확인 창·동의어 사전 창) → 메일머지 화면(tab2.js) 순서로
// 불러와 각 모듈이 자기 화면 요소에 이벤트를 단다.

const common = require('./common.js');
require('./tab2.js');

common.log('프로그램이 준비되었습니다.', 'ok');
