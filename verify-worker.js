// verify-worker.js
// ④ 변환 파일 확인을 별도 프로세스에서 한다. 큰 파일은 다시 여는 데(SheetJS 파싱) 몇 초가 걸리는데,
// 화면 프로세스에서 하면 그동안 창이 멈추고 취소도 못 하기 때문이다. 탭1이 child_process.fork로
// (ELECTRON_RUN_AS_NODE) 띄우고, 메시지 하나를 받아 결과 하나를 돌려준 뒤 끝난다. 취소는 프로세스 종료.
// 받는 것: { filePath, headers, labels, sheets: [{ name, rows }], numberFormats, columnKinds: [{ kind }] }

const fs = require('fs');
const XLSX = require('xlsx');
const writer = require('./excel-writer.js');
const check = require('./output-check.js');

/** 탭1과 같은 계산(같은 함수)으로 확인한다 — 프로세스 안팎 어디서든 결과가 같다 */
function verifyFile({ filePath, headers, labels, sheets, numberFormats, columnKinds }) {
  const workbook = XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellNF: true, dense: true });
  const allRows = sheets.length === 1 ? sheets[0].rows : sheets.flatMap((s) => s.rows);
  const formatOf = writer.makeCellFormatter(headers, allRows, { numberFormats, columnKinds });
  return check.verifyWorkbook(workbook, { headers, labels, sheets, formatOf, toSerial: writer.toExcelDateSerial });
}

if (require.main === module) {
  process.once('message', (job) => {
    let reply;
    try {
      reply = { ok: true, result: verifyFile(job) };
    } catch (err) {
      reply = { ok: false, message: err && err.message ? err.message : String(err) };
    }
    process.send(reply, () => process.exit(0));
  });
}

module.exports = { verifyFile };
