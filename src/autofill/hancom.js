// hancom.js
// 한글(한컴오피스)을 Windows 자동 조작(COM, "HWPFrame.HwpObject")으로 다룬다: 서식 열기·.hwpx/.pdf로 저장·
// 프린터로 바로 인쇄. Node에서 COM을 직접 부르지 않고 PowerShell 스크립트를 만들어 실행한 뒤 결과(JSON)를 읽는다.
// - 한글이 자동 조작으로 파일을 열 때 "접근을 허용할까요?" 창을 띄울 수 있다. 한컴 보안 모듈
//   (FilePathCheckerModule)이 등록되어 있으면 창 없이 진행된다. probe가 등록 여부를 알려준다.
// - 실행기(runner)를 바꿔 끼울 수 있어 Windows가 아닌 곳에서도 스크립트 생성·결과 해석을 단위 테스트한다.

const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const PROG_ID = 'HWPFrame.HwpObject';
const SECURITY_MODULE_KEY = 'HKCU:\\Software\\HNC\\HwpAutomation\\Modules';

/** PowerShell 작은따옴표 문자열 */
const psq = (s) => `'${String(s).replace(/'/g, "''")}'`;

const HEADER = `$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$result = [ordered]@{ ok = $false; steps = @() }
function Add-Step($name, $ok, $detail) { $script:result.steps += [ordered]@{ name = $name; ok = $ok; detail = "$detail" } }
$hwp = $null
`;

const FOOTER = `
} catch {
  $result.error = $_.Exception.Message
} finally {
  if ($hwp -ne $null) { try { [void]$hwp.Clear(1) } catch {} ; try { $hwp.Quit() } catch {} }
}
$result | ConvertTo-Json -Depth 6 -Compress
`;

/** 한글을 띄우고 보안 모듈을 등록하는 공통 앞부분. visible: 한글 창을 보일지 */
function openHwp(visible) {
  return `try {
  $hwp = New-Object -ComObject ${psq(PROG_ID)}
  Add-Step 'create' $true ''
  $mod = $false
  try { $mod = [bool]$hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule') } catch {}
  $result.securityModule = $mod
  try { $hwp.XHwpWindows.Item(0).Visible = ${visible ? '$true' : '$false'} } catch {}
`;
}

const FORMAT_OF = (file) => {
  const ext = path.extname(file).toLowerCase();
  if (ext === '.hwpx' || ext === '.hwpt') return 'HWPX';
  if (ext === '.hwp') return 'HWP';
  if (ext === '.pdf') return 'PDF';
  throw new Error(`한글로 다룰 수 없는 파일 형식입니다(${ext || '확장자 없음'})`);
};

// 형식을 알려 주고 열어 보고, 안 되면 한글이 형식을 스스로 판단하게 해서 한 번 더 연다
const openFile = (file) => `  $opened = [bool]$hwp.Open(${psq(file)}, ${psq(FORMAT_OF(file))}, 'forceopen:true')
  if (-not $opened) { $opened = [bool]$hwp.Open(${psq(file)}, '', 'forceopen:true') }
  Add-Step 'open' $opened ${psq(file)}
  if (-not $opened) { throw '한글이 파일을 열지 못했습니다' }
`;

/** 대화상자 없이 인쇄(부수 1~99, 프린터 이름은 고른 때만) */
const printBlock = (args) => {
  const copies = Math.max(1, Math.min(99, Number(args.copies) || 1));
  return `  $set = $hwp.HParameterSet.HPrint
  [void]$hwp.HAction.GetDefault('Print', $set.HSet)
  $set.NumCopy = ${copies}
  $set.Collate = 1
${args.printer ? `  $set.PrinterName = ${psq(args.printer)}\n` : ''}  $printed = [bool]$hwp.HAction.Execute('Print', $set.HSet)
  Add-Step 'print' $printed 'copies=${copies}'
  if (-not $printed) { throw '한글이 인쇄를 시작하지 못했습니다' }
`;
};

/**
 * 한글의 "모두 바꾸기"로 문서의 글자를 바꾼다(찾을 글자 → 바꿀 글자). 한글이 직접 고치므로 파일 구조가 깨질 일이 없다.
 * 바꾼 뒤 문서 글자에 {{…}}가 남았으면 멈춘다(빈칸인 채로 인쇄하지 않도록). 남은 이름은 leftover로 알린다(값은 남기지 않음).
 */
const replaceBlock = (pairs) => `  $fr = $hwp.HParameterSet.HFindReplace
${pairs
  .map(
    ([find, to]) => `  [void]$hwp.HAction.GetDefault('AllReplace', $fr.HSet)
  $fr.FindString = ${psq(find)}
  $fr.ReplaceString = ${psq(to)}
  try { $fr.IgnoreReplaceString = 0 } catch {}
  try { $fr.FindRegExp = 0 } catch {}
  try { $fr.ReplaceMode = 1 } catch {}
  $fr.IgnoreMessage = 1
  try { $fr.Direction = $hwp.FindDir('AllDoc') } catch {}
  try { $fr.FindType = 1 } catch {}
  [void]$hwp.HAction.Execute('AllReplace', $fr.HSet)
`
  )
  .join('')}  Add-Step 'replace' $true '${pairs.length}'
  $text = [string]$hwp.GetTextFile('TEXT', '')
  $left = @([regex]::Matches($text, '\\{\\{[^{}\\r\\n]{1,30}\\}\\}') | ForEach-Object { $_.Value } | Select-Object -Unique)
  $result.leftover = $left
  if ($left.Count -gt 0) { throw ('서식에 채우지 못한 칸이 남았습니다: ' + ($left -join ', ')) }
`;

const saveAs = (file) => `  $saved = [bool]$hwp.SaveAs(${psq(file)}, ${psq(FORMAT_OF(file))}, '')
  Add-Step 'save' $saved ${psq(file)}
  if (-not $saved) { throw '한글이 파일을 저장하지 못했습니다' }
`;

/**
 * 작업별 PowerShell 스크립트
 * - probe: 한글 연결·버전·보안 모듈 등록 상태
 * - convert: { input, output } — .hwp → .hwpx, 무엇이든 → .pdf
 * - print: { file, copies, printer } — 대화상자 없이 인쇄
 * - printReplace: { file, replacements: [[찾을 글자, 바꿀 글자]], copies, printer, output? } — 원본을 열어 한글이 직접 바꾼 뒤 인쇄(output이면 저장)
 * - selftest: { hwpxOut, pdfOut } — 새 문서에 글자를 넣어 .hwpx·.pdf로 저장했다가 다시 열기
 */
function buildScript(action, args = {}) {
  let body = '';
  if (action === 'probe') {
    body = `${openHwp(false)}  $result.version = "$($hwp.Version)"
  $keys = @()
  if (Test-Path ${psq(SECURITY_MODULE_KEY)}) { (Get-ItemProperty ${psq(SECURITY_MODULE_KEY)}).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' } | ForEach-Object { $keys += "$($_.Name)=$($_.Value)" } }
  $result.registeredModules = $keys
  $result.ok = $true
`;
  } else if (action === 'convert') {
    body = `${openHwp(false)}${openFile(args.input)}${saveAs(args.output)}  $result.ok = Test-Path ${psq(args.output)}
`;
  } else if (action === 'print') {
    body = `${openHwp(false)}${openFile(args.file)}${printBlock(args)}  $result.ok = $true
`;
  } else if (action === 'printReplace') {
    // 원본 서식을 한글로 열어 {{필드}} 글자를 한글이 직접 값으로 바꾸고 인쇄(또는 output으로 저장). 원본 파일은 저장하지 않는다
    const pairs = (args.replacements || []).filter(([find]) => find);
    body = `${openHwp(false)}${openFile(args.file)}${replaceBlock(pairs)}${args.output ? saveAs(args.output) : printBlock(args)}  $result.ok = $true
`;
  } else if (action === 'selftest') {
    body = `${openHwp(false)}  $ins = $hwp.HParameterSet.HInsertText
  [void]$hwp.HAction.GetDefault('InsertText', $ins.HSet)
  $ins.Text = ${psq(args.text || '사회사업실 도우미: 한글 연결 시험 문서')}
  $typed = [bool]$hwp.HAction.Execute('InsertText', $ins.HSet)
  Add-Step 'type' $typed ''
${saveAs(args.hwpxOut)}  [void]$hwp.Clear(1)
${openFile(args.hwpxOut)}${saveAs(args.pdfOut)}  $result.ok = (Test-Path ${psq(args.hwpxOut)}) -and (Test-Path ${psq(args.pdfOut)})
`;
  } else {
    throw new Error(`알 수 없는 작업: ${action}`);
  }
  return HEADER + body + FOOTER;
}

/** 기본 실행기: 스크립트를 UTF-8(BOM) 임시 파일로 써서 PowerShell로 실행. 반환: { code, stdout, stderr, timedOut } */
function powershellRunner(script, { timeoutMs }) {
  return new Promise((resolve) => {
    const file = path.join(os.tmpdir(), `hancom-${process.pid}-${Date.now()}-${Math.random().toString(36).slice(2)}.ps1`);
    fs.writeFileSync(file, `﻿${script}`, 'utf8');
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', file], { windowsHide: true });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    child.stdout.on('data', (d) => (stdout += d));
    child.stderr.on('data', (d) => (stderr += d));
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill();
    }, timeoutMs);
    const done = (code) => {
      clearTimeout(timer);
      try {
        fs.unlinkSync(file);
      } catch (e) {
        /* 임시 스크립트 정리 실패는 무시 */
      }
      resolve({ code, stdout, stderr, timedOut });
    };
    child.on('error', (err) => {
      stderr += String(err && err.message);
      done(-1);
    });
    child.on('close', done);
  });
}

/** PowerShell 출력 마지막 JSON 줄을 결과로. 해석할 수 없으면 실패 결과 */
function parseOutput({ code, stdout, stderr, timedOut }) {
  if (timedOut) return { ok: false, timedOut: true, error: '한글이 제때 응답하지 않았습니다. 한글 창에 "접근 허용" 같은 확인 창이 떠 있는지 확인하세요.' };
  const line = String(stdout || '')
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.startsWith('{'))
    .pop();
  if (!line) return { ok: false, error: `한글 자동 조작 결과를 읽지 못했습니다(종료 코드 ${code}). ${String(stderr || '').trim().slice(0, 300)}` };
  try {
    const r = JSON.parse(line);
    if (!Array.isArray(r.steps)) r.steps = r.steps ? [r.steps] : [];
    if (r.error && /80040154|class not registered|클래스가 등록되지/i.test(r.error)) r.error = '한글(한컴오피스)이 설치되어 있지 않거나 자동 조작 구성 요소가 등록되지 않았습니다.';
    return r;
  } catch (e) {
    return { ok: false, error: `한글 자동 조작 결과 형식이 올바르지 않습니다: ${line.slice(0, 200)}` };
  }
}

/** 작업 실행. options: { runner, timeoutMs, platform } */
async function runHancom(action, args = {}, { runner = powershellRunner, timeoutMs = 90000, platform = process.platform } = {}) {
  if (platform !== 'win32') return { ok: false, error: '한글 자동 조작은 한글이 설치된 Windows에서만 할 수 있습니다.' };
  const script = buildScript(action, args);
  return parseOutput(await runner(script, { timeoutMs }));
}

module.exports = { runHancom, buildScript, parseOutput, powershellRunner, PROG_ID, SECURITY_MODULE_KEY };
