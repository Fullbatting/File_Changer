// 테스트 전용: 프로덕션 preload와 동일한 격리 세계에서 파서를 직접 검증한다.
require('../../preload');
window.__testReadWorkbook = (filePath) => {
  const common = require('../../common');
  const XLSX = require('xlsx');
  const wb = common.readWorkbookSmart(filePath);
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
window.__testFindAutoMatch = (name, columns) => require('../../common').findAutoMatch(name, columns, new Set());
