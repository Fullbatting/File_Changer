// tests/e2e/csv-dates.test.js
// CSV(값이 모두 글자인 원본)의 날짜 열을 실제 날짜로 읽어 저장하는지 확인한다.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/csv-dates.test.js

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { createHarness } = require('./harness.js');
const { sheetFromRows, writeBook, serial } = require('./fixtures.js');

const h = createHarness('csv-dates');
const { js, sleep, click, record } = h;
const cell = (i, b) => js(`const td = document.querySelector('td.wb-cell[data-r="${i}"][data-b="${b}"]'); td ? [td.textContent, td.className] : null`);

async function run() {
  await h.start();
  const csvPath = path.join(h.workDir, '거래내역.csv');
  fs.writeFileSync(
    csvPath,
    '﻿' +
      [
        '코드,거래일,등록시각,혼합,한글날짜,메모',
        '00123,2026-01-05,2026-01-05 13:05,2026-01-05,2026년 1월 5일,첫 거래',
        '00456,2026/2/11,2026-02-11 09:30:15,미정,2026. 2. 11.,',
        '01020,2026.03.20,,2026-03-20,2026년 3월 20일,분기',
      ].join('\r\n') +
      '\r\n'
  );
  const targetPath = writeBook(path.join(h.workDir, '거래양식.xlsx'), { 양식: sheetFromRows([['코드', '거래일', '등록시각', '혼합', '한글날짜', '메모']]) });
  await h.loadFiles(csvPath, targetPath);

  const logText = await js(`document.getElementById('logArea').value`);
  record('날짜 글자 열을 날짜로 읽었다는 로그(거래일·등록시각·한글날짜, 혼합 열은 제외)', logText.includes('날짜로 읽었습니다: 거래일, 등록시각, 한글날짜'), logText.split('\n').slice(-6).join(' / '));

  const d0 = await cell(0, '거래일');
  const d1 = await cell(1, '거래일');
  const d2 = await cell(2, '거래일');
  record('거래일: 형식이 섞여 있어도(2026-01-05 · 2026/2/11 · 2026.03.20) 날짜로 같은 표시', d0[0] === '2026-01-05' && d1[0] === '2026-02-11' && d2[0] === '2026-03-20' && /date/.test(d0[1]), { d0, d1, d2 });
  const t0 = await cell(0, '등록시각');
  record('시각이 있으면 날짜+시각', t0[0] === '2026-01-05 13:05:00' && /date/.test(t0[1]), t0);
  const k1 = await cell(1, '한글날짜');
  record('"2026. 2. 11." · "2026년 1월 5일" 같은 한글 표기도 날짜', k1[0] === '2026-02-11', k1);
  const m0 = await cell(0, '혼합');
  const m1 = await cell(1, '혼합');
  record('날짜가 아닌 값이 섞인 열은 전부 글자 그대로', m0[0] === '2026-01-05' && !/date/.test(m0[1]) && m1[0] === '미정', { m0, m1 });
  record('앞자리 0 코드는 글자 그대로', (await cell(0, '코드'))[0] === '00123', '');

  // 날짜 열에 다른 표기로 고쳐도 날짜로 저장
  await js(`const td = document.querySelector('td.wb-cell[data-r="1"][data-b="거래일"]'); td.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); const i = document.getElementById('wbCellInput'); i.value = '2026.04.01'; i.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); true;`);
  await sleep(150);

  const out = path.join(h.workDir, 'out.xlsx');
  h.queues.save.push(out);
  await click('btnConvert', 300);
  await click('convertReviewProceed', 900);
  const ws = XLSX.readFile(out, { cellNF: true }).Sheets['양식'];
  const at = (r, c) => ws[XLSX.utils.encode_cell({ r, c })] || {};
  record('저장: 거래일은 엑셀 날짜(yyyy-mm-dd)', at(1, 1).t === 'n' && at(1, 1).v === serial(2026, 1, 5) && at(1, 1).z === 'yyyy-mm-dd' && at(3, 1).v === serial(2026, 3, 20), [at(1, 1), at(3, 1)]);
  record('저장: 날짜 열에 "2026.04.01"로 고친 칸도 날짜', at(2, 1).t === 'n' && at(2, 1).v === serial(2026, 4, 1), at(2, 1));
  const expectTime = serial(2026, 1, 5) + (13 * 60 + 5) / 1440;
  record('저장: 등록시각은 날짜+시각(yyyy-mm-dd hh:mm:ss)', at(1, 2).t === 'n' && Math.abs(at(1, 2).v - expectTime) < 1e-6 && at(1, 2).z === 'yyyy-mm-dd hh:mm:ss', at(1, 2));
  record('저장: 혼합 열은 글자', at(1, 3).t === 's' && at(2, 3).v === '미정', [at(1, 3), at(2, 3)]);
  record('저장: 코드는 글자 00123', at(1, 0).t === 's' && at(1, 0).v === '00123', at(1, 0));
}

run()
  .then(() => h.finish())
  .catch((err) => {
    console.log(`ERROR ${err && err.stack ? err.stack : err}`);
    h.finish();
  });
