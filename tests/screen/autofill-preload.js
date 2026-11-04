// 테스트 전용: 프로덕션 preload와 같은 격리 세계에서 서식 자동 기입의 한글 실행기를 가짜 인쇄로 바꾼다.
// 가짜 인쇄는 한글에 넘긴 채운 임시 파일을 printedDir에 복사해 두고 성공을 돌려준다(내용 확인·임시 파일 삭제 확인용).
require('../../src/main/preload');
window.__autofillFakePrint = (printedDir) => {
  const fs = require('fs');
  const path = require('path');
  window.__printedFiles = [];
  require('../../src/renderer/app-shell.js').app.hooks.hancom = async (action, args) => {
    if (action !== 'print') return { ok: false, error: `테스트에서는 ${action}을 하지 않습니다` };
    const copy = path.join(printedDir, `p${window.__printedFiles.length}.hwpx`);
    fs.copyFileSync(args.file, copy);
    window.__printedFiles.push({ temp: args.file, copy, copies: args.copies });
    return { ok: true, steps: [] };
  };
  return true;
};
