// tests/e2e/form-structure.test.js
// 양식 구조 인식: 제목·기준일 줄 아래의 헤더, 맨 아래 합계 행(이름·SUM 수식), 표 아래 설명 줄을 데이터에서 빼고 알리며,
// 이름으로 못 정한 타겟 컬럼을 타겟 양식의 예시 값으로 원본 컬럼에 잇는다(확인 필요). 값은 모두 가짜다.
// 실행: TZ=Asia/Seoul xvfb-run -a node_modules/.bin/electron --no-sandbox --disable-gpu tests/e2e/form-structure.test.js

const { createHarness } = require('./harness.js');
const { makeFormStructure } = require('./fixtures.js');

const h = createHarness('form-structure');
const { js, record } = h;
const q = JSON.stringify;

const recog = (w) => js(`Array.from(document.querySelectorAll('#recog${w} li')).map((li) => (li.className === 'warn' ? '⚠' : '') + li.textContent)`);
const rowInfo = (b) =>
  js(`const r = document.querySelector('.wb-row[data-b=' + ${q(q(b))} + ']');
    r ? { a: r.querySelector('.wb-a-card b') ? r.querySelector('.wb-a-card b').textContent : '', badge: r.querySelector('.wb-b-card .wb-badge').textContent } : null`);
const chip = (id) => js(`document.getElementById(${q(id)}).textContent`);
const logText = () => js(`document.getElementById('logArea').value`);

async function run() {
  await h.start();
  const { source: sourcePath, target: targetPath } = makeFormStructure(h.workDir);
  await h.loadFiles(sourcePath, targetPath);
  record('구조 분석 완료', (await js(`document.getElementById('analyzeStatus').textContent`)) === '분석 완료', await js(`document.getElementById('analyzeStatus').textContent`));

  // ① 인식 결과
  const a = await recog('A');
  record('원본: 헤더 3행 · 데이터 5행(합계·안내 줄 뺌)', a[0] === '헤더: 3행(자동) · 데이터 5행', a);
  record('원본: 합계 행을 행 번호와 함께 알림', a.some((t) => t.startsWith('⚠합계·소계 행 1개는 데이터에서 뺌: 9행')), a);
  record('원본: 표 아래 설명 줄을 알림', a.some((t) => t.startsWith('⚠표 아래 설명 줄 1개는 데이터에서 뺌: 11행')), a);
  const b = await recog('B');
  record('타겟: 헤더 3행, 합계 행은 예시 값에서 뺌', b[0] === '헤더: 3행(자동)' && b.some((t) => t.includes('합계·소계 행 1개는 예시 값에서 뺌: 6행')), b);

  // ② 자동 연결: 이름 + 값
  const info = {};
  for (const col of ['접수일자', '고객 성명', '연락처', '상품', '금액', '부가세', '지점', '비고']) info[col] = await rowInfo(col);
  record('값으로 이음: 접수일자 ← 가입일(값 일치)', info.접수일자.a === '가입일' && info.접수일자.badge === '값 일치', info.접수일자);
  record('값으로 이음: 금액 ← 납입금액(값 일치)', info.금액.a === '납입금액' && info.금액.badge === '값 일치', info.금액);
  record('이름으로 이음은 그대로: 연락처·비고 같은 이름', info.연락처.badge === '같은 이름' && info.비고.badge === '같은 이름', [info.연락처, info.비고]);
  record('원본에 없는 컬럼(부가세·지점)은 연결 없음', info.부가세.badge === '연결 없음' && info.지점.badge === '연결 없음', [info.부가세, info.지점]);
  const fuzzyChip = await chip('wbStatFuzzy');
  record('값 연결도 확인 필요로 셈(비슷한 이름 2 + 값 2)', fuzzyChip === '확인 필요 4', fuzzyChip);
  const log = await logText();
  record('로그: 값으로 이은 근거(예시 값 2개 중 2개)', log.includes('"접수일자" 컬럼을 값이 겹치는 "가입일" 컬럼에 자동 매핑했습니다(타겟 예시 값 2개 중 2개가 같음)'), log.split('\n').slice(-14).join(' / '));

  // 미리보기: 합계 행 없이 5행
  const rowNums = await js(`Array.from(document.querySelectorAll('#wbGrid tbody td.wb-rownum')).map((td) => td.textContent)`);
  record('미리보기: 원본 4~8행만(합계 9행 없음)', JSON.stringify(rowNums) === JSON.stringify(['4', '5', '6', '7', '8']), rowNums);

  // [맞음]으로 확인하면 확인됨
  await js(`document.querySelector('.wb-row[data-b="접수일자"] [data-action="confirm"]').click(); true;`);
  await h.sleep(200);
  const after = await rowInfo('접수일자');
  record('[맞음]: 값 연결도 확인됨으로, 확인 필요 3', after.badge === '확인됨' && (await chip('wbStatFuzzy')) === '확인 필요 3', [after, await chip('wbStatFuzzy')]);
}

run()
  .then(() => h.finish())
  .catch((err) => h.fail(err));
