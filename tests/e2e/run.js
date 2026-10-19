// tests/e2e/run.js
// 매핑 작업대 자동 테스트를 차례로 실행한다: npm run test:e2e
// Linux(화면 없는 서버)에서는 xvfb-run -a npm run test:e2e 로 실행한다. 날짜 검사는 한국 시간 기준이다.

const { spawnSync } = require('child_process');
const path = require('path');
const electronPath = require('electron'); // Node에서 require하면 설치된 Electron 실행 파일 경로

const TESTS = ['visual-toggle.test.js', 'csv-dates.test.js', 'merge-split.test.js', 'workbench.test.js', 'many-columns.test.js', 'performance.test.js'];
const only = process.argv.slice(2);
const args = process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : [];

let failed = 0;
for (const file of only.length > 0 ? only : TESTS) {
  console.log(`\n▶ ${file}`);
  const result = spawnSync(electronPath, [...args, path.join(__dirname, file)], {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...process.env, TZ: process.env.TZ || 'Asia/Seoul' },
    encoding: 'utf-8',
  });
  // Electron 자체의 진단 메시지(dbus·GPU 등)는 빼고 테스트 결과만 보여준다
  (result.stdout || '')
    .split('\n')
    .filter((line) => /^(PASS|FAIL|ERROR|UNHANDLED|\[)/.test(line) && !/dbus|gpu/i.test(line))
    .forEach((line) => console.log(line));
  if (result.status !== 0) failed += 1;
}
console.log(failed === 0 ? '\n모든 테스트 통과' : `\n실패한 테스트 파일 ${failed}개`);
process.exit(failed === 0 ? 0 : 1);
