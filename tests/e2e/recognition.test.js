// tests/e2e/recognition.test.js
// ① 파일 인식: 헤더 행 판정·직접 고르기, 머리글이 빈 열, CSV 인코딩, 1904 날짜, 확장자와 다른 내용.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/recognition.test.js

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const iconv = require('iconv-lite');
const { createHarness } = require('./harness.js');
const { sheetFromRows, writeBook } = require('./fixtures.js');

const h = createHarness('recognition');
const { js, sleep, click, record } = h;

const recog = (w) => js(`Array.from(document.querySelectorAll('#recog${w} li')).map((li) => (li.className === 'warn' ? '⚠' : '') + li.textContent)`);
const choose = async (id, value) => {
  await js(`const el = document.getElementById(${JSON.stringify(id)}); el.value = ${JSON.stringify(value)}; el.dispatchEvent(new Event('change')); true;`);
  await sleep(200);
};
const cellText = (i, b) => js(`const td = document.querySelector('td.wb-cell[data-r="${i}"][data-b="${b}"]'); td ? td.textContent : null`);
const logTail = () => js(`document.getElementById('logArea').value.split('\\n').slice(-12).join(' / ')`);

async function pick(which, filePath) {
  h.queues.open.push([filePath]);
  await click(`drop${which}`, 250);
}

async function run() {
  await h.start();
  const target = writeBook(path.join(h.workDir, '양식.xlsx'), { 양식: sheetFromRows([['이름', '전화', '주소']]) });

  // 1) 머리글 칸 하나가 빈 표: 데이터 행이 더 많이 채워져 있어도 1행이 헤더, C열은 "(이름 없음) C열"
  const src1 = writeBook(path.join(h.workDir, '빈머리글.xlsx'), {
    명단: sheetFromRows([
      ['이름', '전화', ''],
      ['홍길동', '010-1111-2222', '서울'],
      ['김영희', '010-3333-4444', '부산'],
    ]),
  });
  await pick('A', src1);
  await pick('B', target);
  const r1 = await recog('A');
  record('인식 결과: 헤더 1행(자동) · 데이터 2행', r1[0] === '헤더: 1행(자동) · 데이터 2행', r1);
  record('인식 결과: 머리글이 빈 C열을 버리지 않고 알림', r1.some((t) => t.startsWith('⚠') && t.includes('(이름 없음) C열')), r1);
  const options = await js(`Array.from(document.getElementById('headerSelectA').options).map((o) => o.textContent)`);
  record('헤더 행 고르기 목록: 자동 + 맨 위 행 미리보기', options[0] === '자동 (1행)' && options[1] === '1행: 이름, 전화' && options[2].startsWith('2행: 홍길동'), options);
  record('타겟(B)에도 인식 결과(데이터 행 수 없이)', (await recog('B'))[0] === '헤더: 1행(자동)', await recog('B'));
  await click('btnAnalyze', 400);
  const headers = await js(`document.getElementById('logArea').value.includes('원본(A) 컬럼 3개')`);
  record('분석: 원본 컬럼 3개(빈 머리글 열 포함)', headers, await logTail());
  record('미리보기: 주소는 비어 있고 이름·전화 연결', (await cellText(0, '이름')) === '홍길동' && (await cellText(1, '전화')) === '010-3333-4444', [await cellText(0, '이름'), await cellText(1, '전화')]);

  // 2) 헤더 행을 직접 바꾸면 분석이 초기화되고 그 행 기준으로 읽는다
  await choose('headerSelectA', '1');
  const r2 = await recog('A');
  record('헤더 행 직접 고르기: 2행(직접 고름), 위 1행은 뺐다고 알림', r2[0] === '헤더: 2행(직접 고름) · 데이터 1행' && r2.some((t) => t.includes('헤더 위 1개 행') && t.includes('1행 "이름, 전화"')), r2);
  const status = await js(`document.getElementById('analyzeStatus').textContent`);
  record('읽기 설정을 바꾸면 분석 결과 초기화', status === '대기 중', status);
  await choose('headerSelectA', '');
  record('자동으로 되돌리기', (await recog('A'))[0] === '헤더: 1행(자동) · 데이터 2행', await recog('A'));

  // 3) CP949 CSV: 인코딩 자동 판별, 직접 바꾸면 깨진 글자를 알림
  const csvPath = path.join(h.workDir, '명단_cp949.csv');
  fs.writeFileSync(csvPath, iconv.encode('이름,전화,주소\r\n홍길동,010-1111-2222,서울\r\n', 'cp949'));
  await pick('A', csvPath);
  const r3 = await recog('A');
  const encVisible = await js(`!document.getElementById('encodingSelectA').hidden`);
  record('CSV: 인코딩 CP949(자동)으로 읽고 고르기 상자 표시', encVisible && r3.some((t) => t === '글자 인코딩: CP949(한글 윈도우)(자동)'), r3);
  await choose('encodingSelectA', 'utf-8');
  const r4 = await recog('A');
  record('인코딩을 UTF-8로 바꾸면 깨진 글자 경고', r4.some((t) => t.startsWith('⚠') && t.includes('UTF-8(직접 고름)') && t.includes('깨진 글자')), r4);
  await choose('encodingSelectA', 'auto');
  await click('btnAnalyze', 400);
  record('CP949 CSV 분석: 한글 그대로', (await cellText(0, '이름')) === '홍길동' && (await cellText(0, '주소')) === '서울', [await cellText(0, '이름'), await cellText(0, '주소')]);

  // 4) 1904 날짜 체계 엑셀: 날짜가 4년 어긋나지 않음
  const ws = sheetFromRows([['이름', '전화', '주소', '입사일'], ['홍길동', '010', '서울', { t: 'n', v: 46037 - 1462, z: 'yyyy-mm-dd' }]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '명단');
  wb.Workbook = { ...(wb.Workbook || {}), WBProps: { date1904: true } };
  const p1904 = path.join(h.workDir, '맥엑셀_1904.xlsx');
  XLSX.writeFile(wb, p1904);
  const target2 = writeBook(path.join(h.workDir, '양식2.xlsx'), { 양식: sheetFromRows([['이름', '입사일']]) });
  await pick('A', p1904);
  await pick('B', target2);
  record('1904 날짜 체계 알림', (await recog('A')).some((t) => t.includes('1904 날짜 체계')), await recog('A'));
  await click('btnAnalyze', 400);
  record('1904 날짜: 2026-01-15 그대로(2022-01-14 아님)', (await cellText(0, '입사일')) === '2026-01-15', await cellText(0, '입사일'));

  // 5) 내용은 CSV인데 확장자가 .xlsx: 내용에 맞춰 읽고 알림
  const fake = path.join(h.workDir, '이름만바꾼.xlsx');
  fs.writeFileSync(fake, '이름,입사일\r\n홍길동,2026-01-15\r\n');
  await pick('A', fake);
  const r5 = await recog('A');
  record('확장자와 다른 내용 알림', r5.some((t) => t.startsWith('⚠') && t.includes('확장자는 .xlsx이지만 실제 내용은 글자 파일(CSV)')), r5);
  await click('btnAnalyze', 400);
  record('내용에 맞춰 CSV로 읽음(날짜 글자도 날짜로)', (await cellText(0, '이름')) === '홍길동' && (await cellText(0, '입사일')) === '2026-01-15', [await cellText(0, '이름'), await cellText(0, '입사일')]);
}

run()
  .then(() => h.finish())
  .catch((err) => h.fail(err));
