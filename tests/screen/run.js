// tests/screen/run.js
// 화면 시나리오(서식 등록 → 출력 → 인적사항 지움)를 실행한다: npm run test:screen
// Linux(화면 없는 서버)에서는 xvfb-run -a npm run test:screen. 날짜 검사는 한국 시간 기준이다.
// 테스트용 파일은 실행할 때마다 임시 폴더에 만든다(fixtures-*.js). 시나리오마다 PASS/FAIL 줄과
// "N/M passed" 요약 줄을 출력하며, FAIL 줄이 없고 요약이 M/M일 때만 성공으로 본다.

const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const electronPath = require('electron');

const TIMEOUT_MS = 5 * 60 * 1000;
const electronArgs = process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : [];
const env = (fix) => ({ ...process.env, TZ: process.env.TZ || 'Asia/Seoul', SCREEN_FIX: fix });

const basicDir = fs.mkdtempSync(path.join(os.tmpdir(), 'autofill-screen-'));
const made = spawnSync(process.execPath, [path.join(__dirname, 'fixtures-autofill.js')], { env: env(basicDir), encoding: 'utf-8' });
if (made.status !== 0) {
  console.log(`✗ fixtures-autofill.js: 테스트 파일을 만들지 못했습니다\n${made.stderr}`);
  process.exit(1);
}

const RUNS = [{ name: '서식 자동 기입(등록 → 출력 → 인적사항 지움)', file: 'autofill.test.js', dir: basicDir }];

let failed = 0;
for (const run of RUNS) {
  console.log(`\n▶ ${run.name}`);
  const result = spawnSync(electronPath, [...electronArgs, path.join(__dirname, run.file), ...(run.args || [])], {
    env: env(run.dir),
    encoding: 'utf-8',
    timeout: TIMEOUT_MS,
    maxBuffer: 64 * 1024 * 1024,
  });
  const lines = (result.stdout || '').split('\n').filter((l) => /^(PASS|FAIL|UNHANDLED)|passed/.test(l));
  lines.filter((l) => !l.startsWith('PASS')).forEach((l) => console.log(l.trim()));
  const summary = lines.map((l) => /(\d+)\/(\d+) passed/.exec(l)).filter(Boolean).pop();
  const problems = [];
  if (result.error) problems.push(result.error.code === 'ETIMEDOUT' ? '시간 초과' : String(result.error));
  if (result.status !== 0) problems.push(`종료 코드 ${result.status}`);
  if (lines.some((l) => /^(FAIL|UNHANDLED)/.test(l))) problems.push('실패 줄 있음');
  if (!summary) problems.push('결과 요약 줄 없음');
  else if (summary[1] !== summary[2] || summary[2] === '0') problems.push(`요약 ${summary[1]}/${summary[2]}`);
  if (problems.length > 0) {
    failed += 1;
    console.log(`✗ ${run.name}: ${problems.join(', ')}`);
  }
}

fs.rmSync(basicDir, { recursive: true, force: true });
console.log(failed === 0 ? '\n화면 시나리오 모두 통과' : `\n실패한 시나리오 ${failed}개`);
process.exit(failed === 0 ? 0 : 1);
