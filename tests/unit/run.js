// tests/unit/run.js
// 단위 테스트(Electron 없이 Node로)를 차례로 실행한다: npm run test:unit

const { spawnSync } = require('child_process');
const path = require('path');

const TESTS = ['table-reader.test.js', 'auto-match.test.js', 'merge-sources.test.js'];
let failed = 0;
for (const file of TESTS) {
  const result = spawnSync(process.execPath, [path.join(__dirname, file)], { stdio: 'inherit' });
  if (result.status !== 0) {
    failed += 1;
    console.log(`✗ ${file}: 종료 코드 ${result.status}`);
  }
}
console.log(failed === 0 ? '\n단위 테스트 모두 통과' : `\n실패한 단위 테스트 파일 ${failed}개`);
process.exit(failed === 0 ? 0 : 1);
