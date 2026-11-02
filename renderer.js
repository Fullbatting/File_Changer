// renderer.js
// 렌더러 프로세스 진입점: 공통 모듈 → 메일머지 자동기입(tab2.js) 순서로 불러와 각 모듈이
// 자신의 DOM/이벤트를 직접 바인딩하도록 한다. tab2.js가 common.js를 먼저 필요로 하므로
// common.js를 가장 먼저 불러온다.

const common = require('./common.js');
require('./tab2.js');

common.log('프로그램이 준비되었습니다.', 'ok');
