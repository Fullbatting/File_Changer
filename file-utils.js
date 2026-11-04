// file-utils.js
// 파일 읽기·쓰기와 오류 문구, 파일이 바뀌면 저절로 버려지는 캐시. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.

const fs = require('fs');

/** Node.js 파일 시스템 오류 코드를 이해하기 쉬운 한국어 문구로 */
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
      return `파일이 다른 프로그램(Excel/Word 등)에서 열려 있어 사용할 수 없습니다${target}. 해당 프로그램에서 파일을 닫은 뒤 다시 시도하세요.`;
    case 'EISDIR':
      return '선택한 경로는 폴더입니다. 파일을 선택하세요.';
    case 'ENOSPC':
      return '디스크 여유 공간이 부족합니다.';
    default:
      return err.message ? String(err.message) : String(err);
  }
}

/** 파일 읽기. 실패하면 오류에 경로를 붙여 다시 던진다(friendlyErrorMessage가 경로를 보여 준다) */
function safeReadFileSync(filePath, encoding) {
  if (!fs.existsSync(filePath)) {
    const err = new Error(`파일을 찾을 수 없습니다 (${filePath})`);
    err.code = 'ENOENT';
    err.path = filePath;
    throw err;
  }
  try {
    return encoding ? fs.readFileSync(filePath, encoding) : fs.readFileSync(filePath);
  } catch (err) {
    err.path = err.path || filePath;
    throw err;
  }
}

/** 파일 쓰기. 실패하면 오류에 경로를 붙여 다시 던진다 */
function safeWriteFileSync(filePath, data) {
  try {
    fs.writeFileSync(filePath, data);
  } catch (err) {
    err.path = err.path || filePath;
    throw err;
  }
}

/** 파일의 현재 수정 시각(ms). 파일이 없으면 null */
function getFileMtimeMs(filePath) {
  return fs.existsSync(filePath) ? fs.statSync(filePath).mtimeMs : null;
}

/**
 * 파일에서 만든 값(읽은 워크북·템플릿 등)을 담아 두는 캐시. 같은 파일을 단계마다 다시 읽고 해석하지 않게 한다.
 * 파일 수정 시각이 바뀌면(밖에서 다시 저장) 새로 만든다. 여러 파일을 오가도 커지지 않게 limit개까지(먼저 넣은 것부터 버림).
 * get(filePath, load, key = filePath): 담아 둔 값이 없거나 낡았으면 load()로 만든다.
 */
function createFileCache(limit = 8) {
  const entries = new Map(); // key -> { mtimeMs, value }
  return {
    get(filePath, load, key = filePath) {
      const mtimeMs = getFileMtimeMs(filePath);
      const cached = entries.get(key);
      if (cached && cached.mtimeMs === mtimeMs) return cached.value;
      const value = load();
      entries.delete(key);
      if (entries.size >= limit) entries.delete(entries.keys().next().value);
      entries.set(key, { mtimeMs, value });
      return value;
    },
    clear: () => entries.clear(),
    get size() {
      return entries.size;
    },
  };
}

module.exports = { friendlyErrorMessage, safeReadFileSync, safeWriteFileSync, getFileMtimeMs, createFileCache };
