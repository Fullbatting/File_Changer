// settings.js
// 이 PC의 설정: 담당자(이름·직위·번호), 기본 부수, 프린터 이름. 사용자 데이터 폴더의 settings.json에 저장한다.
// 환자 인적사항은 여기 두지 않는다. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.

const fs = require('fs');
const path = require('path');

const FILE_NAME = 'settings.json';
const DEFAULTS = { staffName: '', staffTitle: '', staffPhone: '', copies: 1, printer: '' };
const KEYS = Object.keys(DEFAULTS);

/** 반환: { settings, error } — 파일이 없거나 망가졌으면 기본값(망가진 파일은 그대로 둔다) */
function loadSettings(dir) {
  const file = path.join(dir, FILE_NAME);
  if (!fs.existsSync(file)) return { settings: { ...DEFAULTS }, error: null };
  try {
    return { settings: clean(JSON.parse(fs.readFileSync(file, 'utf8'))), error: null };
  } catch (err) {
    return { settings: { ...DEFAULTS }, error: err };
  }
}

/** 알려진 항목만, 알맞은 모양으로 */
function clean(raw) {
  const s = { ...DEFAULTS };
  KEYS.forEach((k) => {
    if (raw && raw[k] !== undefined) s[k] = raw[k];
  });
  ['staffName', 'staffTitle', 'staffPhone', 'printer'].forEach((k) => {
    s[k] = String(s[k] == null ? '' : s[k]).trim();
  });
  s.copies = Math.max(1, Math.min(99, Math.round(Number(s.copies)) || 1));
  return s;
}

function saveSettings(dir, settings) {
  const s = clean(settings);
  fs.writeFileSync(path.join(dir, FILE_NAME), JSON.stringify(s, null, 2));
  return s;
}

module.exports = { DEFAULTS, loadSettings, saveSettings };
