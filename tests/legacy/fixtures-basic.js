const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const DIR = process.env.LEGACY_FIX;

// --- Tab1 fixtures: A(source) has synonym/exact columns, B(target) has a title row above header ---
const wbA = XLSX.utils.book_new();
const wsA = XLSX.utils.aoa_to_sheet([
  ['이름', '생년월일', '전화번호', '이메일주소'],
  ['홍길동', '1990-01-01', '010-1111-2222', 'hong@example.com'],
  ['김영희', '1992-05-05', '010-3333-4444', 'kim@example.com'],
]);
XLSX.utils.book_append_sheet(wbA, wsA, 'Sheet1');
XLSX.writeFile(wbA, path.join(DIR, 'sourceA.xlsx'));

const wbB = XLSX.utils.book_new();
const wsB = XLSX.utils.aoa_to_sheet([
  ['2026년 명단'], // title row -> header auto-detect should skip this
  ['성명', '생일', '연락처', '이메일', '비고'],
]);
XLSX.utils.book_append_sheet(wbB, wsB, 'Sheet1');
XLSX.writeFile(wbB, path.join(DIR, 'targetB.xlsx'));


