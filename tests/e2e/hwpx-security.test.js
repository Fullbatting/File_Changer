const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const { createHarness } = require('./harness');

const h = createHarness('hwpx-security');
const { js, click, sleep, record } = h;

async function run() {
  await h.start();
  const marker = path.join(h.workDir, 'executed');
  const attack = `<img src="missing" onerror="require('fs').writeFileSync('${marker}','owned')"/>`;
  const encoded = '&lt;svg onload="alert(1)"/&gt;';
  // 메일머지 템플릿은 XML로 해석한다: <hp:t> 안의 <img>는 글자가 아닌 요소라 변수 이름이 될 수 없고(문서에 있던 그대로 남음),
  // &lt;svg…&gt;는 글자 "<svg…>"로 풀린 변수 이름이 된다. 어느 쪽도 화면에서 요소로 그려지거나 실행되면 안 된다.
  const decoded = '<svg onload="alert(1)"/>';
  const template = path.join(h.workDir, 'unsafe.hwpx');
  const zip = new PizZip();
  const xml = `<hs:sec xmlns:hs="urn:section" xmlns:hp="urn:paragraph"><hp:p><hp:run><hp:t>안녕 {{이</hp:t></hp:run><hp:run><hp:t>름}} {{${attack}}} {{${encoded}}}</hp:t></hp:run></hp:p></hs:sec>`;
  zip.file('Contents/section0.xml', xml);
  fs.writeFileSync(template, zip.generate({ type: 'nodebuffer' }));

  record('공격 픽스처는 올바른 XML', await js(`!new DOMParser().parseFromString(${JSON.stringify(xml)}, 'application/xml').querySelector('parsererror')`));

  const pageGlobals = await h.win.webContents.executeJavaScript(`({
    require: typeof require, process: typeof process, ipc: typeof window.ipcRenderer
  })`);
  record('웹 페이지에서 Node와 IPC가 노출되지 않음',
    pageGlobals.require === 'undefined' && pageGlobals.process === 'undefined' && pageGlobals.ipc === 'undefined', pageGlobals);

  h.queues.open.push([template]);
  await click('dropTemplate', 200);
  await click('btnExtractVars', 200);
  await sleep(300);
  const form = await js(`({
    status: document.getElementById('extractStatus').textContent,
    labels: Array.from(document.querySelectorAll('#varFormGrid .tag')).map(e => e.textContent),
    inputNames: Array.from(document.querySelectorAll('#varFormGrid input')).map(e => e.dataset.varName),
    images: document.querySelectorAll('#varFormGrid img').length
  })`);
  record('악성 글자 변수는 요소가 아닌 텍스트, 요소는 변수가 아님, 일반 변수도 추출',
    form.status === '2개 변수 발견' && form.labels.join('|') === `{{이름}}|{{${decoded}}}` &&
    form.inputNames.join('|') === `이름|${decoded}` && form.images === 0 && !fs.existsSync(marker), form);

  await js(`for (const input of document.querySelectorAll('#varFormGrid input')) {
    input.value = input.dataset.varName === '이름' ? '홍길동 & 가족' : '안전한 값';
  } true;`);
  const manual = path.join(h.workDir, 'manual.hwpx');
  h.queues.save.push(manual);
  await click('btnGenerateDocx', 400);
  const manualXml = fs.existsSync(manual) ? new PizZip(fs.readFileSync(manual)).file('Contents/section0.xml').asText() : '';
  record('직접 입력 및 저장 IPC: 나뉜 변수와 악성 이름 변수를 정상 치환',
    manualXml.includes('홍길동 &amp; 가족') && manualXml.includes('안전한 값') && !manualXml.includes('<svg') && !manualXml.includes('{{이') && !fs.existsSync(marker), manualXml);

  const csv = path.join(h.workDir, 'data.csv');
  fs.writeFileSync(csv, '\uFEFF이름\n김영희\n');
  await js(`const r = document.querySelector('input[name="mmMode"][value="batch"]'); r.checked = true; r.dispatchEvent(new Event('change', { bubbles: true })); true;`);
  h.queues.open.push([csv]);
  await click('dropMailMergeData', 200);
  await click('btnLoadMailMergeData', 300);
  const mapping = await js(`({
    names: Array.from(document.querySelectorAll('#mailMergeMappingTableBody tr td:first-child')).map(e => e.textContent),
    selected: Array.from(document.querySelectorAll('#mailMergeMappingTableBody select')).map(e => [e.dataset.templateVar, e.value]),
    images: document.querySelectorAll('#mailMergeMappingTableBody img').length
  })`);
  record('일괄 매핑에서 변수 이름은 텍스트이고 일반 변수 자동 매핑',
    mapping.names.includes(`{{${decoded}}}`) && mapping.selected.some(([name, value]) => name === '이름' && value === '이름') && mapping.images === 0, mapping);
  await click('btnGenerateBatch', 200);
  const review = await js(`({
    open: !document.getElementById('mailMergeReviewModal').hidden,
    head: Array.from(document.querySelectorAll('#mailMergeReviewTableHead th')).map(e => e.textContent),
    summary: document.getElementById('mailMergeReviewSummary').textContent,
    images: document.querySelectorAll('#mailMergeReviewModal img').length
  })`);
  record('일괄 미리보기 제목과 경고에서도 악성 변수는 텍스트',
    review.open && review.head.includes('이름') && review.head.includes(decoded) && review.summary.includes(decoded) &&
    review.images === 0 && !fs.existsSync(marker), review);
  const outDir = path.join(h.workDir, 'batch');
  fs.mkdirSync(outDir);
  h.queues.folder.push(outDir);
  await click('mailMergeReviewProceed', 400);
  const files = fs.readdirSync(outDir);
  const batchXml = files.length ? new PizZip(fs.readFileSync(path.join(outDir, files[0]))).file('Contents/section0.xml').asText() : '';
  record('일괄 생성 및 폴더 IPC: 정상 변수 치환, 코드 미실행',
    files.length === 1 && batchXml.includes('김영희') && !batchXml.includes('<svg') && !fs.existsSync(marker), batchXml);

  const otherFile = require('url').pathToFileURL(path.join(__dirname, '..', '..', 'styles.css')).href;
  await h.win.webContents.executeJavaScript(`window.open('about:blank'); location.href = ${JSON.stringify(otherFile)}; true;`);
  await sleep(200);
  record('팝업과 파일 페이지 이탈이 차단됨',
    h.win.webContents.getURL().endsWith('/index.html') && require('electron').BrowserWindow.getAllWindows().length === 1);
  await h.win.webContents.executeJavaScript(`location.href = 'about:blank'; true;`);
  await sleep(200);
  const blank = await h.win.webContents.executeJavaScript(`({ url: location.href, require: typeof require, process: typeof process })`);
  record('특수 페이지 이동에도 Node를 노출하거나 앱을 초기화하지 않음',
    blank.url === 'about:blank' && blank.require === 'undefined' && blank.process === 'undefined', blank);
}

run().then(() => h.finish()).catch((err) => h.fail(err));
