// tests/e2e/fixtures.js
// 테스트용 엑셀 파일을 만든다(저장소에 바이너리를 두지 않기 위해 실행할 때마다 임시 폴더에 생성).

const path = require('path');
const XLSX = require('xlsx');

/** 날짜를 시간대 영향 없이 엑셀 일련번호로 */
function serial(y, m, d) {
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

/** 값 배열(행 목록)을 origin 위치부터 시트로 만든다. 값이 { t, v, z } 객체면 셀 그대로 쓴다 */
function sheetFromRows(rows, origin = { r: 0, c: 0 }) {
  const ws = {};
  let maxR = origin.r;
  let maxC = origin.c;
  rows.forEach((row, ri) => {
    row.forEach((value, ci) => {
      if (value === '' || value === null || value === undefined) return;
      const r = origin.r + ri;
      const c = origin.c + ci;
      const cell = typeof value === 'object' ? value : typeof value === 'number' ? { t: 'n', v: value } : { t: 's', v: String(value) };
      ws[XLSX.utils.encode_cell({ r, c })] = cell;
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    });
  });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  return ws;
}

function writeBook(filePath, sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, ws]) => XLSX.utils.book_append_sheet(wb, ws, name));
  XLSX.writeFile(wb, filePath);
  return filePath;
}

const dateCell = (y, m, d) => ({ t: 'n', v: serial(y, m, d), z: 'yyyy-mm-dd' });
const moneyCell = (v) => ({ t: 'n', v, z: '#,##0' });

/**
 * 작은 예제: 원본은 B2에서 시작하고 제목 행이 있는 표(헤더 3행, 데이터 4~7행).
 * 7행은 5행과 모든 값이 같다(타겟으로 옮기면 중복 행). "빈칸컬럼"은 값이 하나도 없다.
 */
function makeSmall(dir) {
  const source = sheetFromRows(
    [
      ['2026 직원 명부'],
      ['이름', '사번', '입사일', '급여', '부서', '휴대폰', '우편번호', '비고', '빈칸컬럼'],
      ['홍길동', '20230001', dateCell(2026, 1, 15), moneyCell(1500000), '영업팀', '010-1234-5678', '03187', '첫째', ''],
      ['김영희', '20230002', dateCell(2026, 2, 19), moneyCell(2250000), '영업팀', '010-9876-5432', '04321', '', ''],
      ['이철수', '20230003', dateCell(2026, 3, 2), moneyCell(1980000), '개발팀', '', '06236', '셋째', ''],
      ['김영희', '20230002', dateCell(2026, 2, 19), moneyCell(2250000), '영업팀', '010-9876-5432', '04321', '', ''],
    ],
    { r: 1, c: 1 }
  );
  const target = sheetFromRows([['성명', '사번', '입사릴', '급여', '부서명', '연락처']]);
  return {
    source: writeBook(path.join(dir, '직원명부.xlsx'), { 직원: source }),
    target: writeBook(path.join(dir, '회사양식.xlsx'), { 명부: target }),
  };
}

/** 원본 52열 × 타겟 41열. 같은 이름 28 · 동의어 5 · 비슷한 이름 4 · 연결 없음 4(그중 "주소"는 후보 2개가 동률) */
const MANY_FIELDS = Array.from({ length: 27 }, (_, i) => `필드${String(i + 1).padStart(2, '0')}`);
const MANY_TARGET = [
  '성명', '사번', '생일', '이메일', '소속', '직위',
  ...MANY_FIELDS,
  '입사일자', '직책수당', '주민번호', '계좌번호',
  '연락처', '국민연금', '주소', '비고',
];
const MANY_SOURCE = [
  '이름', '사번', '생년월일', '이메일주소', '부서', '직급',
  ...MANY_FIELDS,
  '입사일', '직급수당', '주민등록번호', '급여계좌번호',
  '휴대폰', '비상연락처', '국민연금보험료', '주소1', '주소2',
  '빈칸1', '빈칸2', '빈칸3', '빈칸4', '빈칸5', '빈칸6',
  '우편번호', '학력', '병역', 'E-mail',
];

function makeMany(dir, rowCount = 120) {
  const rows = [MANY_SOURCE];
  for (let i = 0; i < rowCount; i++) {
    const n = i + 1;
    const values = {
      이름: `직원${n}`,
      사번: `2019${String(n).padStart(4, '0')}`,
      생년월일: dateCell(1990, (i % 12) + 1, (i % 27) + 1),
      이메일주소: `user${n}@corp.kr`,
      부서: ['영업팀', '개발팀', '인사팀'][i % 3],
      직급: ['사원', '대리', '과장'][i % 3],
      입사일: dateCell(2019 + (i % 6), (i % 12) + 1, 1),
      직급수당: moneyCell((i % 3) * 150000),
      주민등록번호: `9${String(n).padStart(5, '0')}-1******`,
      급여계좌번호: `110-234-${String(100000 + n)}`,
      휴대폰: `010-${String(1000 + n).slice(-4)}-${String(5000 + n).slice(-4)}`,
      비상연락처: `010-2222-${String(3000 + n).slice(-4)}`,
      국민연금보험료: moneyCell(135000 + n * 100),
      주소1: '서울 종로구',
      주소2: `${n}번지`,
      우편번호: String(3000 + n).padStart(5, '0'),
      학력: ['대졸', '석사'][i % 2],
      병역: ['군필', '면제'][i % 2],
      'E-mail': `p${n}@corp.kr`,
    };
    MANY_FIELDS.forEach((f, k) => {
      values[f] = `${f}-${n}-${k}`;
    });
    rows.push(MANY_SOURCE.map((h) => (h.startsWith('빈칸') ? '' : values[h])));
  }
  return {
    source: writeBook(path.join(dir, '인사기록_전체.xlsx'), { 인사기록: sheetFromRows(rows) }),
    target: writeBook(path.join(dir, '급여대장_양식.xlsx'), { 급여대장: sheetFromRows([MANY_TARGET]) }),
  };
}

/** 성능 측정용: 같은 이름 cols개 컬럼 × rowCount행 (원본/타겟 열 수를 따로 줄 수 있다) */
function makeWide(dir, { rowCount, sourceCols, targetCols, name }) {
  const sourceHeaders = Array.from({ length: sourceCols }, (_, i) => `컬럼${String(i + 1).padStart(3, '0')}`);
  const targetHeaders = sourceHeaders.slice(0, targetCols);
  const rows = [sourceHeaders];
  for (let r = 0; r < rowCount; r++) {
    rows.push(sourceHeaders.map((h, c) => (c % 4 === 0 ? r * 10 + c : `값${r}-${c}`)));
  }
  return {
    source: writeBook(path.join(dir, `${name}_source.xlsx`), { 데이터: sheetFromRows(rows) }),
    target: writeBook(path.join(dir, `${name}_target.xlsx`), { 양식: sheetFromRows([targetHeaders]) }),
  };
}

module.exports = { makeSmall, makeMany, makeWide, MANY_TARGET, MANY_SOURCE, serial, sheetFromRows, writeBook, dateCell, moneyCell };
