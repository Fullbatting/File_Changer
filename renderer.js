// renderer.js
// 렌더러 프로세스 진입점: 공통 모듈 → 탭1(문서 양식 변환기) → 탭2(메일머지) 순서로
// 불러와 각 모듈이 자신의 DOM/이벤트를 직접 바인딩하도록 한다. 탭1/탭2는 서로를
// require하지 않고 common.js에만 의존하므로 두 모듈 사이의 순서는 영향이 없지만,
// 둘 다 common.js를 먼저 필요로 하므로 common.js를 가장 먼저 불러온다.

const common = require('./common.js');
require('./tab1.js');
require('./tab2.js');

common.log('프로그램이 준비되었습니다.', 'ok');
