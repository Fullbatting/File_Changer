// tests/unit/launch.test.js
// 시작 프로그램(src/launcher/launch.js): 버전 확인, 준비가 필요한지, 프로그램이 "떴다"고 알릴 때까지 기다리기.
// 프로그램 대신 가짜 프로세스로 확인한다. node tests/unit/launch.test.js

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { EventEmitter } = require('events');
const launch = require('../../src/launcher/launch.js');

let passed = 0;
let failed = 0;
const queue = [];
const test = (title, fn) => queue.push({ title, fn });
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'launch-unit-'));

test('Node.js 버전: 22.12 이상만', () => {
  assert.ok(launch.nodeVersionOk('22.12.0') && launch.nodeVersionOk('v24.1.0') && launch.nodeVersionOk('22.20.1'));
  assert.ok(!launch.nodeVersionOk('22.11.9') && !launch.nodeVersionOk('20.18.0') && !launch.nodeVersionOk('v18.0.0'));
});

test('준비가 필요한지: 처음 · 업데이트(package.json이 준비 때와 다름) · 실행 엔진 없음, 모두 같으면 바로 실행', () => {
  const ok = { pkg: '{"a":1}', marker: '{"a":1}', hasModules: true, electronPath: '/x/electron' };
  assert.deepStrictEqual(launch.setupReasons(ok), []);
  assert.strictEqual(launch.setupReasons({ ...ok, hasModules: false }).length, 1);
  assert.ok(launch.setupReasons({ ...ok, marker: null })[0].includes('업데이트'));
  assert.ok(launch.setupReasons({ ...ok, pkg: '{"a":2}' })[0].includes('업데이트'));
  assert.ok(launch.setupReasons({ ...ok, electronPath: null })[0].includes('실행 엔진'));
});

/** 가짜 프로그램: spawn을 흉내 내고, 받은 환경 변수의 알림 파일에 결과를 쓰거나 끝난다 */
function fakeSpawn(behave) {
  const calls = [];
  const spawn = (exe, args, opts) => {
    const child = new EventEmitter();
    child.unref = () => (child.unrefd = true);
    calls.push({ exe, args, opts, child });
    setTimeout(() => behave(child, opts.env[launch.READY_ENV]), 50);
    return child;
  };
  return { spawn, calls };
}

test('프로그램이 "떴다"고 알리면 성공, 따로 떨어진 프로세스(detached·unref)로 띄우고 알림 파일은 지운다', async () => {
  const readyFile = path.join(dir, 'ready1.json');
  const f = fakeSpawn((child, file) => fs.writeFileSync(file, JSON.stringify({ ok: true })));
  const r = await launch.startApp('/x/electron', { spawn: f.spawn, readyFile, timeoutMs: 3000 });
  assert.deepStrictEqual(r, { ok: true });
  const { opts, child } = f.calls[0];
  assert.ok(opts.detached && opts.stdio === 'ignore' && child.unrefd);
  assert.strictEqual(opts.env[launch.READY_ENV], readyFile);
  assert.ok(!fs.existsSync(readyFile));
});

test('프로그램이 실패를 알리면(이미 오류 창을 띄움) 그 내용 그대로', async () => {
  const f = fakeSpawn((child, file) => fs.writeFileSync(file, JSON.stringify({ ok: false, error: '화면 준비 실패', shown: true })));
  const r = await launch.startApp('/x/electron', { spawn: f.spawn, readyFile: path.join(dir, 'ready2.json'), timeoutMs: 3000 });
  assert.deepStrictEqual(r, { ok: false, error: '화면 준비 실패', shown: true });
});

test('알림 없이 닫히면 · 제때 안 뜨면 · 실행 파일을 못 띄우면 실패(시작 프로그램이 오류 창을 띄움)', async () => {
  const closed = await launch.startApp('/x/electron', { spawn: fakeSpawn((child) => child.emit('exit', 1)).spawn, readyFile: path.join(dir, 'ready3.json'), timeoutMs: 3000 });
  assert.ok(!closed.ok && !closed.shown && closed.error.includes('닫혔습니다(종료 코드 1)'), closed.error);
  const slow = await launch.startApp('/x/electron', { spawn: fakeSpawn(() => {}).spawn, readyFile: path.join(dir, 'ready4.json'), timeoutMs: 300 });
  assert.ok(!slow.ok && slow.error.includes('제때 열리지 않았습니다'), slow.error);
  const missing = await launch.startApp('/x/electron', { spawn: fakeSpawn((child) => child.emit('error', new Error('ENOENT'))).spawn, readyFile: path.join(dir, 'ready5.json'), timeoutMs: 3000 });
  assert.ok(!missing.ok && missing.error.includes('ENOENT'), missing.error);
});

(async () => {
  for (const t of queue) {
    try {
      await t.fn();
      passed += 1;
      console.log(`PASS - ${t.title}`);
    } catch (err) {
      failed += 1;
      console.log(`FAIL - ${t.title} :: ${err.message}`);
    }
  }
  fs.rmSync(dir, { recursive: true, force: true });
  console.log(`\n[launch] ${passed}/${passed + failed} 통과`);
  process.exit(failed ? 1 : 0);
})();
