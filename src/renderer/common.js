// common.js
// 화면 공통 도우미: 진행 기록(로그 — 접으면 마지막 한 줄만, 펼치면 전체)과 오류 알림. 진행 기록에는 인적사항 값을 남기지 않는다(부르는 쪽이 지킨다).

const logArea = document.getElementById('logArea');
const logLast = document.getElementById('logLast');
const logToggle = document.getElementById('btnToggleLog');
let logScrollPending = false;

function log(message, level = 'info') {
  const time = new Date().toLocaleTimeString('ko-KR', { hour12: false });
  const prefix = { info: 'INFO', ok: ' OK ', warn: 'WARN', error: 'FAIL' }[level] || 'INFO';
  // value +=는 로그 전체를 다시 배치하므로 끝에 덧붙이기만 한다
  const end = logArea.value.length;
  logArea.setRangeText(`[${time}] [${prefix}] ${message}\n`, end, end, 'end');
  // 접어 두었을 때는 마지막 한 줄만 보인다
  logLast.textContent = `${time}  ${message}`;
  logLast.className = `fa-log-last ${level}`;
  // 맨 아래로 스크롤은 다음 화면 갱신 때 한 번만(바로 scrollHeight를 읽으면 강제 레이아웃이 생긴다)
  if (logScrollPending) return;
  logScrollPending = true;
  requestAnimationFrame(() => {
    logScrollPending = false;
    logArea.scrollTop = logArea.scrollHeight;
  });
}

document.getElementById('btnClearLog').addEventListener('click', () => {
  logArea.value = '';
  logLast.textContent = '';
});

logToggle.addEventListener('click', () => {
  const open = logArea.hidden;
  logArea.hidden = !open;
  logToggle.setAttribute('aria-expanded', String(open));
  if (open) logArea.scrollTop = logArea.scrollHeight;
});

/** 파일 오류 코드를 알기 쉬운 문구로 */
function friendlyErrorMessage(err) {
  if (!err) return '알 수 없는 오류입니다.';
  const target = err.path ? ` (${err.path})` : '';
  switch (err.code) {
    case 'ENOENT':
      return `파일을 찾을 수 없습니다${target}. 파일이 이동되었거나 삭제되지 않았는지 확인하세요.`;
    case 'EACCES':
    case 'EPERM':
      return `파일에 접근할 권한이 없습니다${target}. 관리자 권한 또는 파일 권한을 확인하세요.`;
    case 'EBUSY':
      return `파일이 다른 프로그램(한글 등)에서 열려 있어 사용할 수 없습니다${target}. 그 프로그램에서 파일을 닫은 뒤 다시 시도하세요.`;
    case 'EISDIR':
      return '선택한 경로는 폴더입니다. 파일을 선택하세요.';
    case 'ENOSPC':
      return '디스크 여유 공간이 부족합니다.';
    default:
      return err.message ? String(err.message) : String(err);
  }
}

function notifyError(context, err) {
  const msg = friendlyErrorMessage(err);
  log(`${context}: ${msg}`, 'error');
  alert(`❌ ${context}\n\n${msg}`);
}

module.exports = { log, notifyError, friendlyErrorMessage };
