// tests/legacy/run.js
// 예전 회귀 시나리오(탭1·탭2 기본 흐름, 한글 HWPX 표·템플릿, 전체 기능 점검 3단계)를 실행한다: npm run test:legacy
// Linux(화면 없는 서버)에서는 xvfb-run -a npm run test:legacy. 날짜 검사는 한국 시간 기준이다.
// 테스트용 파일은 실행할 때마다 임시 폴더에 만든다(fixtures-*.js). 시나리오마다 PASS/FAIL 줄과
// "N/M passed" 요약 줄을 출력하며, FAIL 줄이 없고 요약이 M/M일 때만 성공으로 본다.

const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const electronPath = require('electron');

const TIMEOUT_MS = 5 * 60 * 1000;
const electronArgs = process.platform === 'linux' ? ['--no-sandbox', '--disable-gpu'] : [];
const env = (fix) => ({ ...process.env, TZ: process.env.TZ || 'Asia/Seoul', LEGACY_FIX: fix });

// 기본·HWPX 시나리오가 함께 쓰는 폴더와 전체 기능 점검 폴더를 따로 둔다(파일 이름이 대소문자만 다른 것이 있음)
const basicDir = fs.mkdtempSync(path.join(os.tmpdir(), 'legacy-basic-'));
const auditDir = fs.mkdtempSync(path.join(os.tmpdir(), 'legacy-audit-'));
for (const [script, dir] of [['fixtures-basic.js', basicDir], ['fixtures-hwpx.js', basicDir], ['fixtures-hwpx-template.js', basicDir], ['fixtures-audit.js', auditDir]]) {
  const r = spawnSync(process.execPath, [path.join(__dirname, script)], { env: env(dir), encoding: 'utf-8' });
  if (r.status !== 0) {
    console.log(`✗ ${script}: 테스트 파일을 만들지 못했습니다\n${r.stderr}`);
    process.exit(1);
  }
}

const RUNS = [
  { name: 'basic', file: 'basic.test.js', dir: basicDir },
  { name: 'hwpx', file: 'hwpx.test.js', dir: basicDir },
  { name: 'hwpx-template', file: 'hwpx-template.test.js', dir: basicDir },
  { name: 'audit 1(전체 기능)', file: 'audit.test.js', dir: auditDir, args: ['1'] },
  { name: 'audit 2(다시 실행 — 설정 유지)', file: 'audit.test.js', dir: auditDir, args: ['2'] },
  {
    name: 'audit 3(망가진 설정 파일 — 복구)',
    file: 'audit.test.js',
    dir: auditDir,
    args: ['3'],
    before: () => {
      fs.writeFileSync(path.join(auditDir, 'userdata', 'mapping-presets.json'), '{broken');
      fs.writeFileSync(path.join(auditDir, 'userdata', 'synonym-groups.json'), '[[broken');
    },
  },
];

let failed = 0;
for (const run of RUNS) {
  console.log(`\n▶ ${run.name}`);
  if (run.before) run.before();
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
fs.rmSync(auditDir, { recursive: true, force: true });
console.log(failed === 0 ? '\n예전 회귀 시나리오 모두 통과' : `\n실패한 시나리오 ${failed}개`);
process.exit(failed === 0 ? 0 : 1);
