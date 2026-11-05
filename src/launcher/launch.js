// src/launcher/launch.js
// run.bat이 부르는 시작 프로그램: 준비(처음·업데이트 때만) → 프로그램 실행 → 프로그램 화면이 뜬 것을 확인하면 끝낸다.
// 이것이 끝나면 run.bat의 검은 창(cmd)도 닫힌다. 프로그램은 따로 떨어진 프로세스로 계속 돈다.
// 문제가 생기면(오래된 Node.js, 설치 실패, 프로그램이 뜨지 않음) 오류 창으로 알린다.
// Node.js 자체는 run.bat이 먼저 찾거나 내려받아 둔다(src/launcher/node-env.bat).
// PC에 없는 것 가운데 패키지와 실행 엔진은 여기서 내려받는다. 한글(한컴오피스)은 유료 프로그램이라
// 내려받을 수 없으므로, 처음 준비할 때 없으면 설치하라고 안내만 한다.
//
// 오래된 Node.js에서도 "버전이 낮다"는 안내까지는 해야 하므로 최신 문법을 쓰지 않는다.

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const childProcess = require('child_process');

const ROOT = path.resolve(__dirname, '..', '..');
const TITLE = '사회사업실 도우미';
const MIN_NODE = [22, 12];
const READY_TIMEOUT_MS = 90 * 1000;
// 준비가 끝났을 때의 package.json 사본. 지금 것과 같으면 다음 실행부터 준비를 건너뛴다
const MARKER = path.join(ROOT, 'node_modules', '.ready-package.json');
const READY_ENV = 'SWH_READY_FILE';

/** "22.12.0" 같은 버전이 최소 버전 이상인지 */
function nodeVersionOk(version, min) {
  const want = min || MIN_NODE;
  const parts = String(version).replace(/^v/, '').split('.').map(Number);
  return parts[0] > want[0] || (parts[0] === want[0] && parts[1] >= want[1]);
}

/**
 * 준비가 필요한 까닭들(빈 배열이면 바로 실행). state: { pkg, marker, hasModules, electronPath }
 * - 처음(패키지 없음) · 업데이트(package.json이 준비 때와 다름) · 실행 엔진 없음
 */
function setupReasons(state) {
  const reasons = [];
  if (!state.hasModules) reasons.push('처음 실행이라 필요한 패키지를 설치합니다.');
  else if (state.pkg !== state.marker) reasons.push('업데이트로 필요한 패키지가 바뀌어 다시 설치합니다.');
  if (!state.electronPath) reasons.push('프로그램 실행 엔진(Electron)을 내려받습니다.');
  return reasons;
}

/** 프로그램이 남긴 시작 결과(JSON 한 줄)를 읽는다. 아직 없으면 null */
function readReady(file) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (e) {
    return null;
  }
}

const readText = (file) => {
  try {
    return fs.readFileSync(file, 'utf8');
  } catch (e) {
    return null;
  }
};

/** 설치된 실행 엔진(electron 실행 파일) 경로, 없으면 null */
function electronPath() {
  try {
    const pathFile = path.join(ROOT, 'node_modules', 'electron', 'path.txt');
    const exe = path.join(ROOT, 'node_modules', 'electron', 'dist', fs.readFileSync(pathFile, 'utf8').trim());
    return fs.existsSync(exe) ? exe : null;
  } catch (e) {
    return null;
  }
}

/** 메시지 상자(Windows). kind: 'Error' | 'Information'. 그 밖의 운영체제에서는 화면 글자만 */
function showBox(message, title, kind) {
  if (process.platform !== 'win32') return;
  childProcess.spawnSync(
    'powershell.exe',
    ['-NoProfile', '-Command', "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show($env:SWH_MESSAGE, $env:SWH_TITLE, 'OK', $env:SWH_KIND)"],
    { env: Object.assign({}, process.env, { SWH_MESSAGE: message, SWH_TITLE: title, SWH_KIND: kind }), windowsHide: true }
  );
}

/** 오류 창 */
function showError(message) {
  console.error(`\n[오류] ${message}\n`);
  showBox(message, `${TITLE} - 오류`, 'Error');
}

/** 안내 창 */
function showNotice(message) {
  console.log(`\n[안내] ${message}\n`);
  showBox(message, TITLE, 'Information');
}

/** 한글(한컴오피스)이 설치되어 있는지(자동 조작용 등록 이름이 있는지). Windows가 아니면 null */
function hancomInstalled() {
  if (process.platform !== 'win32') return null;
  // 32비트 한글도 있으므로 64비트·32비트 등록을 모두 본다
  return ['/reg:64', '/reg:32'].some((view) => childProcess.spawnSync('reg.exe', ['query', 'HKCR\\HWPFrame.HwpObject', view], { stdio: 'ignore', windowsHide: true }).status === 0);
}

/** 명령을 실행하고 출력은 이 창에 그대로 보인다. 반환: 성공 여부 */
function runVisible(command) {
  const r = childProcess.spawnSync(command, { cwd: ROOT, stdio: 'inherit', shell: true });
  return r.status === 0;
}

/** 준비: 패키지 설치 → 실행 엔진 내려받기 → 준비 완료 표시. 실패하면 오류 문구를 던진다 */
function setup(reasons) {
  reasons.forEach((r) => console.log(`- ${r}`));
  console.log('\n[1/2] 필요한 패키지를 설치합니다. 인터넷 연결이 필요하며 몇 분 걸릴 수 있습니다.\n');
  if (!runVisible('npm install')) {
    if (!fs.existsSync(path.join(ROOT, 'node_modules'))) throw new Error('패키지 설치에 실패했습니다. 인터넷 연결 상태를 확인한 뒤 run.bat을 다시 실행하세요.');
    console.log('\n[경고] 패키지 확인 중 오류가 있었습니다. 인터넷 연결 문제일 수 있어 기존 설치 상태로 계속합니다.');
  }
  // Electron 42부터는 패키지 설치 때 실행 파일을 내려받지 않으므로 직접 받는다(이미 있으면 바로 끝난다)
  console.log('\n[2/2] 프로그램 실행 엔진(Electron)을 확인합니다. 처음에는 내려받느라 몇 분 걸릴 수 있습니다.\n');
  // run.bat이 고른 Node.js(지금 이 프로그램을 돌리는 것)로 그대로 실행한다
  if (!runVisible(`"${process.execPath}" node_modules/electron/install.js`) || !electronPath()) {
    throw new Error('프로그램 실행 엔진을 내려받지 못했습니다. 인터넷 연결 상태를 확인한 뒤 run.bat을 다시 실행하세요.');
  }
  fs.copyFileSync(path.join(ROOT, 'package.json'), MARKER);
}

/**
 * 프로그램을 따로 떨어진 프로세스로 띄우고, 화면이 떴다는 알림(ready 파일)을 기다린다.
 * 반환: Promise<{ ok, error?, shown? }> — shown: 프로그램이 이미 오류 창을 띄웠는지
 */
function startApp(exe, { timeoutMs = READY_TIMEOUT_MS, spawn = childProcess.spawn, readyFile } = {}) {
  const file = readyFile || path.join(os.tmpdir(), `swh-ready-${process.pid}-${Date.now()}.json`);
  try {
    fs.rmSync(file, { force: true });
  } catch (e) {
    /* 없으면 그만 */
  }
  return new Promise((resolve) => {
    let done = false;
    const finish = (result) => {
      if (done) return;
      done = true;
      clearInterval(poll);
      clearTimeout(timer);
      try {
        fs.rmSync(file, { force: true });
      } catch (e) {
        /* 임시 파일 정리 실패는 무시 */
      }
      resolve(result);
    };
    const env = Object.assign({}, process.env, { [READY_ENV]: file });
    delete env.ELECTRON_RUN_AS_NODE;
    let child;
    try {
      child = spawn(exe, [ROOT], { cwd: ROOT, env, detached: true, stdio: 'ignore', windowsHide: false });
    } catch (err) {
      resolve({ ok: false, error: `프로그램을 시작하지 못했습니다: ${err.message}` });
      return;
    }
    child.on('error', (err) => finish({ ok: false, error: `프로그램을 시작하지 못했습니다: ${err.message}` }));
    child.on('exit', (code) => finish(readReady(file) || { ok: false, error: `프로그램이 열리자마자 닫혔습니다(종료 코드 ${code}). run.bat을 다시 실행해 보세요.` }));
    if (child.unref) child.unref();
    const poll = setInterval(() => {
      const r = readReady(file);
      if (r) finish(r);
    }, 200);
    const timer = setTimeout(() => finish({ ok: false, error: '프로그램 화면이 제때 열리지 않았습니다. 잠시 뒤 run.bat을 다시 실행해 보세요.' }), timeoutMs);
  });
}

async function main() {
  console.log(`${TITLE}을(를) 시작합니다...`);
  if (!nodeVersionOk(process.versions.node)) {
    showError(`설치된 Node.js 버전(${process.versions.node})이 너무 낮습니다. ${MIN_NODE.join('.')} 이상이 필요합니다.\nhttps://nodejs.org/ 에서 최신 LTS 버전을 설치한 뒤 run.bat을 다시 실행하세요.`);
    return 1;
  }
  const reasons = setupReasons({
    pkg: readText(path.join(ROOT, 'package.json')),
    marker: readText(MARKER),
    hasModules: fs.existsSync(path.join(ROOT, 'node_modules')),
    electronPath: electronPath(),
  });
  if (reasons.length) {
    try {
      setup(reasons);
    } catch (err) {
      showError(err.message);
      return 1;
    }
  }
  const result = await startApp(electronPath());
  if (result.ok) {
    // 처음 준비했을 때만 한 번 알린다(매번 띄우면 귀찮다). 한글이 없어도 서식 등록·입력은 된다
    if (reasons.length && hancomInstalled() === false) {
      showNotice('이 PC에 한글(한컴오피스)이 없습니다.\n한글은 유료 프로그램이라 자동으로 설치할 수 없습니다.\n서류를 출력하거나 한글 파일로 저장하려면 한글을 설치한 뒤 프로그램을 다시 여세요.');
    }
    return 0;
  }
  // 프로그램이 스스로 오류 창을 띄운 경우(화면 준비 실패 등)에는 또 띄우지 않는다
  if (!result.shown) showError(result.error || '프로그램을 시작하지 못했습니다.');
  return 1;
}

if (require.main === module) {
  main().then(
    (code) => process.exit(code),
    (err) => {
      showError(`시작하는 중 예상하지 못한 오류가 발생했습니다: ${err && err.message ? err.message : err}`);
      process.exit(1);
    }
  );
}

module.exports = { nodeVersionOk, setupReasons, readReady, startApp, READY_ENV, MIN_NODE };
