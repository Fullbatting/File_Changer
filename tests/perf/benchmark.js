// tests/perf/benchmark.js
// 대표 파일 크기별로 "파일 읽기부터" 단계별 시간과 메모리를 잰다: npm run bench
// 화면(그리기·작업대) 없이 앱과 같은 모듈로 읽기 → 표 만들기 → 자동 연결 → 엑셀 쓰기(스트리밍) →
// 다시 열어 확인까지. 크기마다 따로 프로세스를 띄워 메모리가 섞이지 않게 한다.
// 결과는 표로 출력하고, --json이면 JSON 한 줄씩.

const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const CASES = [
  { name: 'xlsx 1천행×20열', rows: 1000, cols: 20, type: 'xlsx' },
  { name: 'xlsx 1만행×20열', rows: 10000, cols: 20, type: 'xlsx' },
  { name: 'xlsx 5만행×20열', rows: 50000, cols: 20, type: 'xlsx' },
  { name: 'xlsx 1만행×100열', rows: 10000, cols: 100, type: 'xlsx' },
  { name: 'csv 5만행×20열', rows: 50000, cols: 20, type: 'csv' },
];

function makeFile(dir, { rows, cols, type }) {
  const XLSX = require('xlsx');
  const headers = Array.from({ length: cols }, (_, i) => `컬럼${String(i + 1).padStart(3, '0')}`);
  const value = (r, c) => {
    switch (c % 5) {
      case 0: return r * 10 + c; // 숫자
      case 1: return `직원${r}`; // 글자
      case 2: return String(r % 1000).padStart(5, '0'); // 앞자리 0 코드
      case 3: return type === 'csv' ? `2026-01-${String((r % 28) + 1).padStart(2, '0')}` : { t: 'n', v: 46037 + (r % 365), z: 'yyyy-mm-dd' }; // 날짜
      default: return ['영업팀', '개발팀', '인사팀'][r % 3];
    }
  };
  const file = path.join(dir, `bench_${rows}x${cols}.${type}`);
  if (type === 'csv') {
    const lines = [headers.join(',')];
    for (let r = 0; r < rows; r++) lines.push(headers.map((_, c) => value(r, c)).join(','));
    fs.writeFileSync(file, `﻿${lines.join('\r\n')}\r\n`);
  } else {
    const ws = {};
    headers.forEach((h, c) => (ws[XLSX.utils.encode_cell({ r: 0, c })] = { t: 's', v: h }));
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const v = value(r, c);
        ws[XLSX.utils.encode_cell({ r: r + 1, c })] = typeof v === 'object' ? v : typeof v === 'number' ? { t: 'n', v } : { t: 's', v };
      }
    }
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows, c: cols - 1 } });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '데이터');
    XLSX.writeFile(wb, file, { compression: true });
  }
  return { file, headers };
}

/** 자식 프로세스: 한 파일을 처음부터 끝까지 처리하며 단계별 시간·메모리를 잰다 */
async function runCase(file, cols) {
  const XLSX = require('xlsx');
  const reader = require('../../table-reader.js');
  const autoMatch = require('../../auto-match.js');
  const writer = require('../../excel-writer.js');
  const check = require('../../output-check.js');
  const MB = 1024 * 1024;
  let peakRss = 0;
  const sample = () => {
    peakRss = Math.max(peakRss, process.memoryUsage().rss);
  };
  const timer = setInterval(sample, 20); // 비동기 단계(쓰기) 중에도 재기
  const stages = [];
  const stage = async (name, fn) => {
    global.gc && global.gc();
    sample();
    const t0 = process.hrtime.bigint();
    const out = await fn();
    const ms = Number(process.hrtime.bigint() - t0) / 1e6;
    sample();
    stages.push({ name, ms, rssMB: process.memoryUsage().rss / MB });
    return out;
  };

  const size = fs.statSync(file).size;
  const ext = path.extname(file);
  const wb = await stage('파일 읽기', () => {
    const buffer = fs.readFileSync(file);
    if (ext === '.csv') return XLSX.read(reader.decodeTextBuffer(buffer).text, { type: 'string', raw: true, dense: true });
    return XLSX.read(buffer, { type: 'buffer', cellNF: true, dense: true });
  });
  const table = await stage('표 만들기', () => reader.buildTableFromSheet(wb.Sheets[wb.SheetNames[0]], wb.SheetNames[0], { date1904: reader.isDate1904(wb) }));
  const targetHeaders = table.headers.slice().reverse();
  const matches = await stage('자동 연결', () => autoMatch.autoMatchAll(targetHeaders, table.headers, []));
  const rows = await stage('결과 행 만들기', () => table.rows.map((r) => Object.fromEntries(targetHeaders.map((b) => [b, matches[b].match ? r[matches[b].match] : '']))));
  const out = path.join(path.dirname(file), `out_${path.basename(file, ext)}.xlsx`);
  const spec = { headers: targetHeaders, labels: targetHeaders, sheets: [{ name: '결과', rows }] };
  const formatOf = writer.makeCellFormatter(targetHeaders, rows, {});
  await stage('엑셀 쓰기(스트리밍)', () => writer.writeStyledExcel({ ...spec, formatOf }, out));
  const result = await stage('다시 열어 확인', () => {
    const reread = XLSX.read(fs.readFileSync(out), { type: 'buffer', cellNF: true, dense: true });
    return check.verifyWorkbook(reread, { ...spec, formatOf, toSerial: writer.toExcelDateSerial });
  });
  clearInterval(timer);
  sample();
  const outSize = fs.statSync(out).size;
  fs.unlinkSync(out);
  return { sizeMB: size / MB, outMB: outSize / MB, rows: table.rows.length, cols, ok: result.ok, problems: result.problemCount, stages, peakRssMB: peakRss / MB };
}

if (process.argv[2] === '--child') {
  runCase(process.argv[3], Number(process.argv[4]))
    .then((r) => {
      process.stdout.write(`RESULT ${JSON.stringify(r)}\n`);
      process.exit(0);
    })
    .catch((err) => {
      console.error(err);
      process.exit(1);
    });
} else {
  const json = process.argv.includes('--json');
  const only = process.argv.slice(2).filter((a) => !a.startsWith('--'));
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bench-'));
  const lines = [];
  console.log(`측정 환경: Node ${process.version}, ${os.cpus()[0].model} × ${os.cpus().length}, 메모리 ${(os.totalmem() / 1024 ** 3).toFixed(1)}GB`);
  for (const c of CASES.filter((x) => only.length === 0 || only.some((o) => x.name.includes(o)))) {
    const { file } = makeFile(dir, c);
    const res = spawnSync(process.execPath, ['--expose-gc', '--max-old-space-size=4096', __filename, '--child', file, String(c.cols)], { encoding: 'utf-8', maxBuffer: 16 * 1024 * 1024 });
    const line = (res.stdout || '').split('\n').find((l) => l.startsWith('RESULT '));
    if (!line) {
      console.log(`✗ ${c.name}: 실패 ${res.stderr}`);
      continue;
    }
    const r = JSON.parse(line.slice(7));
    if (json) {
      console.log(JSON.stringify({ name: c.name, ...r }));
      continue;
    }
    const total = r.stages.reduce((s, x) => s + x.ms, 0);
    lines.push(`| ${c.name} | ${r.sizeMB.toFixed(1)}MB | ${r.stages.map((s) => (s.ms / 1000).toFixed(2)).join(' | ')} | ${(total / 1000).toFixed(2)} | ${r.peakRssMB.toFixed(0)}MB | ${r.ok ? '같음' : `다름 ${r.problems}`} |`);
  }
  if (!json) {
    console.log('\n| 파일 | 크기 | 파일 읽기 | 표 만들기 | 자동 연결 | 결과 행 | 엑셀 쓰기 | 다시 열어 확인 | 합계(초) | 최대 메모리 | 확인 |');
    console.log('|---|---|---|---|---|---|---|---|---|---|---|');
    lines.forEach((l) => console.log(l));
  }
  fs.rmSync(dir, { recursive: true, force: true });
}
