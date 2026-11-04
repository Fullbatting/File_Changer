// synonyms.js
// 동의어 사전: "성명"↔"이름"처럼 표기는 달라도 같은 뜻인 이름을 템플릿 변수 ↔ 데이터 열 자동 매핑에서 바로 잇는다.
// 사용자 데이터 폴더의 synonym-groups.json에 저장하고, 화면에서는 "한 줄 = 한 그룹(쉼표 구분)" 글로 고친다.
// 화면(DOM)에 기대지 않아 node로 단위 테스트한다.

const fs = require('fs');
const path = require('path');
const { safeReadFileSync, safeWriteFileSync } = require('./file-utils.js');

const FILE_NAME = 'synonym-groups.json';

const DEFAULT_SYNONYM_GROUPS = [
  ['이름', '성명', '성함'],
  ['생년월일', '출생일자', '생일'],
  ['전화번호', '연락처', '핸드폰번호', '휴대폰번호', '휴대전화'],
  ['이메일', '이메일주소', '메일주소', '전자우편'],
  ['주소', '소재지', '거주지', '주소지'],
  ['부서', '소속', '소속부서', '부서명'],
  ['직급', '직책', '직위'],
  ['나이', '연령'],
  ['회사명', '업체명', '상호', '회사'],
  ['담당자', '담당자명', '책임자'],
  ['성별', 'sex', 'gender'],
  ['비고', '메모', '참고사항'],
];

const defaults = () => DEFAULT_SYNONYM_GROUPS.map((g) => [...g]);
const filePathIn = (dir) => path.join(dir, FILE_NAME);

/**
 * 저장해 둔 사전을 읽는다. 파일이 없으면 기본 사전을 저장하고 그것을 쓴다.
 * 반환: { groups, error } — 파일이 망가졌으면 기본 사전과 error(무엇이 잘못됐는지)를 돌려준다(파일은 그대로 둔다)
 */
function loadSynonymGroups(dir) {
  const file = filePathIn(dir);
  try {
    if (!fs.existsSync(file)) {
      const groups = defaults();
      saveSynonymGroups(dir, groups);
      return { groups, error: null };
    }
    const parsed = JSON.parse(safeReadFileSync(file, 'utf-8'));
    return { groups: Array.isArray(parsed) ? parsed.filter((g) => Array.isArray(g) && g.length >= 2) : defaults(), error: null };
  } catch (err) {
    return { groups: defaults(), error: err };
  }
}

function saveSynonymGroups(dir, groups) {
  safeWriteFileSync(filePathIn(dir), JSON.stringify(groups, null, 2));
}

/** 그룹 배열 → "한 줄 = 한 그룹(쉼표 구분)" 글 */
const groupsToText = (groups) => groups.map((g) => g.join(', ')).join('\n');

/** 글 → 그룹 배열(빈 이름은 빼고, 이름이 둘 이상인 줄만) */
function textToGroups(text) {
  return text
    .split('\n')
    .map((line) =>
      line
        .split(',')
        .map((term) => term.trim())
        .filter((term) => term !== '')
    )
    .filter((group) => group.length >= 2);
}

module.exports = { DEFAULT_SYNONYM_GROUPS, loadSynonymGroups, saveSynonymGroups, groupsToText, textToGroups };
