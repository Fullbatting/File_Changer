// renderer.js
// 격리된 preload 세계의 앱 진입점(preload.js가 index.html에서만 불러온다): 공통 모듈 → 문서 양식 변환기(tab1.js) 순서로 불러와 각 모듈이
// 자신의 DOM/이벤트를 직접 바인딩하도록 한다. tab1.js가 common.js를 먼저 필요로 하므로
// common.js를 가장 먼저 불러온다.

const common = require('./common.js');
require('./tab1.js');

common.log('프로그램이 준비되었습니다.', 'ok');
