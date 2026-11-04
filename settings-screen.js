// settings-screen.js
// ③ 설정 화면: 담당자(이름·직위·번호), 기본 부수, 프린터 이름, 한글 연결 확인.

const { saveSettings } = require('./settings.js');
const { log, notifyError } = require('./common.js');

function init(app) {
  const $ = (id) => document.getElementById(id);
  const fields = { staffName: 'setStaffName', staffTitle: 'setStaffTitle', staffPhone: 'setStaffPhone', copies: 'setCopies', printer: 'setPrinter' };
  const fill = () => Object.entries(fields).forEach(([k, id]) => ($(id).value = app.settings[k]));
  fill();

  $('btnSaveSettings').addEventListener('click', () => {
    try {
      app.settings = saveSettings(app.userDataPath, Object.fromEntries(Object.entries(fields).map(([k, id]) => [k, $(id).value])));
      fill();
      $('settingsSaved').textContent = '저장했습니다';
      setTimeout(() => ($('settingsSaved').textContent = ''), 2500);
      log('설정을 저장했습니다.', 'ok');
      app.changed('settings');
    } catch (err) {
      notifyError('설정을 저장하지 못했습니다', err);
    }
  });

  $('btnProbeHancom').addEventListener('click', async () => {
    $('probeResult').textContent = '확인 중…';
    const r = await app.hancom('probe', {});
    $('probeResult').textContent = r.ok
      ? `한글 연결됨 (버전 ${r.version || '알 수 없음'}) · 보안 모듈 ${r.securityModule ? '등록됨' : '등록 안 됨 — 파일을 열 때 "접근 허용" 창이 뜰 수 있음'}`
      : `연결 안 됨: ${r.error}`;
  });
}

module.exports = { init };
