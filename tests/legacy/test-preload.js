// 테스트 전용: 프로덕션 preload와 동일한 격리 세계에서 파서·자동 매핑을 직접 검증한다.
require('../../preload');
window.__testReadWorkbook = (filePath) => {
  const { readWorkbook } = require('../../data-file');
  const XLSX = require('xlsx');
  const wb = readWorkbook(filePath);
  const sheets = {};
  wb.SheetNames.forEach((name) => {
    sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
  });
  return {
    names: wb.SheetNames,
    preferred: wb.preferredSheetName,
    sheets,
    approvalObjects: XLSX.utils.sheet_to_json(wb.Sheets['표 1'], { defval: '' }),
  };
};
window.__testAutoMatch = (name, columns) => require('../../common').autoMatchAll([name], columns)[name];
