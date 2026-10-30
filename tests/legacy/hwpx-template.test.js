const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const FIX = process.env.LEGACY_FIX;
const out = path.join(FIX, 'template_result.hwpx');
if (fs.existsSync(out)) fs.unlinkSync(out);
const q = [[path.join(FIX, 'template.hwpx')]];
ipcMain.handle('dialog:openFile', async () => q.shift() || []);
ipcMain.handle('dialog:saveFile', async () => out);
ipcMain.handle('app:getUserDataPath', () => fs.mkdtempSync(path.join(os.tmpdir(), 'tpl-')));
ipcMain.handle('dialog:selectFolder', async () => null);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
app.whenReady().then(async () => {
  const win = new BrowserWindow({ show: false, webPreferences: { nodeIntegration: true, contextIsolation: false } });
  await win.loadFile(path.join(__dirname, '..', '..', 'index.html'));
  const js = (c) => win.webContents.executeJavaScript(c);
  await js('window.alert = (m) => { window.__a = m; }; true;');
  await sleep(300);
  await js(`document.querySelector('.tab-btn[data-tab="tab2"]').click(); document.getElementById('dropTemplate').click(); true;`);
  await sleep(300);
  await js(`document.getElementById('btnExtractVars').click(); true;`);
  await sleep(200);
  const status = await js(`document.getElementById('extractStatus').textContent`);
  await js(`document.querySelector('#varFormGrid input').value = '홍길동'; document.getElementById('btnGenerateDocx').click(); true;`);
  await sleep(500);
  const PizZip = require('pizzip');
  const result = fs.existsSync(out) ? new PizZip(fs.readFileSync(out)).file('Contents/section0.xml').asText() : 'NO OUTPUT: ' + (await js('window.__a || ""'));
  const checks = [
    ['한글 템플릿: 두 run으로 나뉜 {{이름}} 변수 1개 추출', status === '1개 변수 발견', status],
    ['한글 템플릿: 나뉜 변수 자리에 값이 들어가고 서식(run)은 유지', result.includes('<hp:t>안녕하세요 홍길동</hp:t>') && result.includes('<hp:t>님</hp:t>'), result],
  ];
  checks.forEach(([name, ok, detail]) => console.log(`${ok ? 'PASS' : 'FAIL'} - ${name}${ok ? '' : ` :: ${detail}`}`));
  const passed = checks.filter((c) => c[1]).length;
  console.log(`=== ${passed}/${checks.length} passed ===`);
  app.exit(passed === checks.length ? 0 : 1);
});
