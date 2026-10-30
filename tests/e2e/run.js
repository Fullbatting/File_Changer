// tests/e2e/run.js
// 매핑 작업대 자동 테스트를 차례로 실행한다: npm run test:e2e
// Linux(화면 없는 서버)에서는 xvfb-run -a npm run test:e2e 로 실행한다. 날짜 검사는 한국 시간 기준이다.

const { spawnSync } = require('child_process');
const path = require('path');
const electronPath = require('electron'); // Node에서 require하면 설치된 Electron 실행 파일 경로

const TESTS = ['visual-toggle.test.js', 'recognition.test.js', 'csv-dates.test.js', 'merge-split.test.js', 'workbench.test.js', 'many-columns.test.js', 'performance.test.js'];
const only = process.argv.slice(2);
const TIMEOUT_MS = 10 * 60 * 1000;
const args = process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : [];

let failed = 0;
for (const file of only.length > 0 ? only : TESTS) {
  console.log(`\n▶ ${file}`);
  const result = spawnSync(electronPath, [...args, path.join(__dirname, file)], {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...process.env, TZ: process.env.TZ || 'Asia/Seoul' },
    encoding: 'utf-8',
    timeout: TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  // Electron 자체의 진단 메시지(dbus·GPU 등)는 빼고 테스트 결과만 보여준다
  const lines = (result.stdout || '').split('\n').filter((line) => /^(PASS|FAIL|ERROR|UNHANDLED|\[)/.test(line) && !/dbus|gpu/i.test(line));
  lines.forEach((line) => console.log(line));
  // 성공은 세 가지가 모두 맞을 때만: 종료 코드 0, FAIL 줄 없음, "[이름] N/N 통과" 요약 줄이 있고 N이 같음
  const summary = lines.map((l) => /^\[[^\]]+\] (\d+)\/(\d+) 통과$/.exec(l)).find(Boolean);
  const problems = [];
  if (result.error) problems.push(result.error.code === 'ETIMEDOUT' ? `시간 초과(${TIMEOUT_MS / 60000}분)` : String(result.error));
  if (result.status !== 0) problems.push(`종료 코드 ${result.status}`);
  if (lines.some((l) => /^(FAIL|ERROR|UNHANDLED)/.test(l))) problems.push('실패 줄 있음');
  if (!summary) problems.push('결과 요약 줄 없음');
  else if (summary[1] !== summary[2] || summary[2] === '0') problems.push(`요약 ${summary[1]}/${summary[2]}`);
  if (problems.length > 0) {
    failed += 1;
    console.log(`✗ ${file}: ${problems.join(', ')}`);
  }
}
console.log(failed === 0 ? '\n모든 테스트 통과' : `\n실패한 테스트 파일 ${failed}개`);
process.exit(failed === 0 ? 0 : 1);
