// tests/unit/hancom.test.js
// 한글 자동 조작(hancom.js) 단위 테스트: Windows·한글 없이 PowerShell 스크립트 생성과 결과 해석을 확인한다.
// 실제 한글 연결은 Windows PC에서 tools\한글연결점검.bat로 확인한다. node tests/unit/hancom.test.js

const assert = require('assert');
const { buildScript, parseOutput, runHancom } = require('../../src/autofill/hancom.js');

let passed = 0;
let failed = 0;
const queue = [];
const test = (title, fn) => queue.push({ title, fn });

/** 따옴표 안을 빼고 괄호 짝이 맞는지 */
function balanced(script) {
  const stripped = script.replace(/'(?:[^']|'')*'/g, "''").replace(/"[^"]*"/g, '""');
  const count = (ch) => (stripped.match(new RegExp(`\\${ch}`, 'g')) || []).length;
  return count('{') === count('}') && count('(') === count(')');
}

test('스크립트: 작업마다 한글 생성 → 보안 모듈 → 작업 → 정리(Clear·Quit) → JSON 출력', () => {
  const scripts = {
    probe: buildScript('probe'),
    convert: buildScript('convert', { input: 'C:\\a\\서식.hwp', output: 'C:\\a\\서식.hwpx' }),
    print: buildScript('print', { file: 'C:\\a\\서식.hwpx', copies: 2 }),
    selftest: buildScript('selftest', { hwpxOut: 'C:\\a\\t.hwpx', pdfOut: 'C:\\a\\t.pdf' }),
  };
  Object.entries(scripts).forEach(([name, s]) => {
    assert.ok(balanced(s), `${name}: 괄호 짝`);
    assert.ok(s.includes("New-Object -ComObject 'HWPFrame.HwpObject'"), name);
    assert.ok(s.includes("RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')"), name);
    assert.ok(/finally \{[\s\S]*Clear\(1\)[\s\S]*Quit\(\)/.test(s), `${name}: 끝나면 문서 버리고 한글 종료`);
    assert.ok(s.trim().endsWith('ConvertTo-Json -Depth 6 -Compress'), name);
  });
  assert.ok(scripts.convert.includes("$hwp.Open('C:\\a\\서식.hwp', 'HWP', 'forceopen:true')"));
  assert.ok(scripts.convert.includes("$hwp.SaveAs('C:\\a\\서식.hwpx', 'HWPX', '')"));
  assert.ok(scripts.print.includes('$set.NumCopy = 2') && scripts.print.includes("HAction.Execute('Print'"));
  assert.ok(!scripts.print.includes('PrinterName'), '프린터를 안 고르면 기본 프린터');
  assert.ok(scripts.selftest.includes("SaveAs('C:\\a\\t.pdf', 'PDF', '')"));
});

test('스크립트: 경로의 작은따옴표는 두 번 써서 안전하게, 부수는 1~99로, 모르는 형식·작업은 거부', () => {
  const s = buildScript('print', { file: "C:\\O'Neil\\서식.hwpx", copies: 500, printer: "3층 'A' 프린터" });
  assert.ok(s.includes("'C:\\O''Neil\\서식.hwpx'"));
  assert.ok(s.includes("$set.PrinterName = '3층 ''A'' 프린터'"));
  assert.ok(s.includes('$set.NumCopy = 99'));
  assert.ok(balanced(s));
  assert.throws(() => buildScript('convert', { input: 'a.docx', output: 'b.hwpx' }), /다룰 수 없는 파일 형식/);
  assert.throws(() => buildScript('format-c'), /알 수 없는 작업/);
});

test('결과 해석: 마지막 JSON 줄, 단계 하나도 배열로, 시간 초과·한글 없음 안내', () => {
  const ok = parseOutput({ code: 0, stdout: '경고 줄\r\n{"ok":true,"steps":{"name":"create","ok":true,"detail":""},"version":"12, 0, 0, 0"}\r\n', stderr: '' });
  assert.strictEqual(ok.ok, true);
  assert.deepStrictEqual(ok.steps.map((x) => x.name), ['create']);
  const timeout = parseOutput({ timedOut: true });
  assert.ok(!timeout.ok && timeout.error.includes('접근 허용'));
  const noHwp = parseOutput({ code: 0, stdout: '{"ok":false,"steps":[],"error":"Retrieving the COM class factory ... 80040154 Class not registered"}' });
  assert.ok(noHwp.error.includes('설치되어 있지 않거나'));
  const garbage = parseOutput({ code: 1, stdout: '', stderr: 'powershell 오류' });
  assert.ok(!garbage.ok && garbage.error.includes('종료 코드 1'));
});

test('실행: Windows가 아니면 실행하지 않고 안내, 실행기는 바꿔 끼울 수 있음', async () => {
  const linux = await runHancom('probe', {}, { platform: 'linux' });
  assert.ok(!linux.ok && linux.error.includes('Windows'));
  let seen = null;
  const r = await runHancom('convert', { input: 'a.hwp', output: 'a.hwpx' }, {
    platform: 'win32',
    runner: async (script, opts) => {
      seen = { script, opts };
      return { code: 0, stdout: '{"ok":true,"steps":[]}', stderr: '' };
    },
    timeoutMs: 1234,
  });
  assert.ok(r.ok);
  assert.strictEqual(seen.opts.timeoutMs, 1234);
  assert.ok(seen.script.includes("'a.hwpx', 'HWPX'"));
});

(async () => {
  for (const { title, fn } of queue) {
    try {
      await fn();
      passed += 1;
      console.log(`PASS - ${title}`);
    } catch (err) {
      failed += 1;
      console.log(`FAIL - ${title} :: ${err.message}`);
    }
  }
  console.log(`\n[hancom] ${passed}/${passed + failed} 통과`);
  process.exit(failed ? 1 : 0);
})();
