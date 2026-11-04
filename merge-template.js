// merge-template.js
// 메일머지 템플릿: Word(.docx/.dotx)·한글(.hwpx/.hwpt) 문서의 {{변수}}를 찾고 값으로 바꾼 결과 문서를 만든다.
// 일괄 생성의 결과 파일 이름·빈 값 검사도 여기 둔다. 화면(DOM)에 기대지 않아 node로 단위 테스트한다.
// - Word: docxtemplater(본문·머리글·바닥글·각주까지 바꾼다)
// - 한글: 서식 엔진(form-model.js의 원문 위치 모델 + form-fill.js의 원문 XML 편집). 여러 런으로 나뉜 {{변수}}도
//   찾고, 바꾼 자리 밖의 XML은 한 글자도 건드리지 않는다.

const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { safeReadFileSync, createFileCache } = require('./file-utils.js');
const { modelFromHwpx } = require('./form-model.js');
const { recognizePlaceholders } = require('./form-recognize.js');
const { writePlan } = require('./form-fill.js');
const { formatPreviewValue } = require('./table-reader.js');

const PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;
const DOTX_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml';
const DOCX_MAIN_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

/** 확장자로 템플릿 형식('docx'|'hwpx')을 정한다. 서식 파일(.dotx/.hwpt)은 본문 형식과 구조가 같다 */
function templateFormat(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.docx' || ext === '.dotx') return 'docx';
  if (ext === '.hwpx' || ext === '.hwpt') return 'hwpx';
  if (ext === '.hwp') {
    throw new Error('구형 .hwp 문서는 메일머지 템플릿으로 쓸 수 없습니다(일괄 생성의 데이터 파일로는 읽을 수 있습니다). 한글 프로그램에서 "다른 이름으로 저장 > HWPX"로 저장한 뒤 그 파일을 템플릿으로 선택하세요.');
  }
  throw new Error('지원하지 않는 템플릿 형식입니다. .docx, .dotx, .hwpx 또는 .hwpt 파일을 선택하세요.');
}

/**
 * docxtemplater의 제어 태그({{#반복}}, {{/반복}}, {{^조건}}, {{.}})는 같은 {{ }}를 쓰지만 변수가 아니다.
 * 반복·조건 화면은 없으므로 입력칸·매핑표에 나오지 않게 뺀다.
 */
const isControlToken = (name) => /^[#/^.]/.test(name.trim());

/** 이름 목록에서 제어 태그를 빼고, 처음 나온 순서대로 한 번씩 */
const uniqueVariables = (names) => [...new Set(names.map((n) => n.trim()).filter((n) => n && !isControlToken(n)))];

/* ---------- Word ---------- */

/** docx 안에서 변수를 찾을 XML: 본문 → 머리글 → 바닥글 → 각주·미주 순(docxtemplater가 모두 바꾸므로 모두 본다) */
function docxParts(zip) {
  const rank = (name) => (name === 'word/document.xml' ? 0 : /^word\/header\d*\.xml$/.test(name) ? 1 : /^word\/footer\d*\.xml$/.test(name) ? 2 : 3);
  return Object.keys(zip.files)
    .filter((name) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name))
    .sort((a, b) => rank(a) - rank(b) || a.localeCompare(b, undefined, { numeric: true }));
}

function docxVariables(buffer) {
  const zip = new PizZip(buffer);
  if (!zip.file('word/document.xml')) throw new Error('올바른 .docx 파일이 아닙니다. (word/document.xml 없음)');
  const names = [];
  docxParts(zip).forEach((name) => {
    // Word는 글자를 여러 <w:t>로 쪼개는 일이 많아, 태그를 지운 글자에서 찾는다
    const plain = zip.file(name).asText().replace(/<[^>]+>/g, '');
    for (const m of plain.matchAll(PLACEHOLDER)) names.push(m[1]);
  });
  return uniqueVariables(names);
}

/** docxtemplater는 render 뒤 내부 상태가 바뀌어 다시 쓸 수 없으므로, 원본 바이트에서 매번 새로 만든다 */
function renderDocx(buffer, data) {
  const doc = new Docxtemplater(new PizZip(buffer), {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
    // 값이 없는 변수는 docxtemplater 기본값("undefined" 글자) 대신 빈칸으로
    nullGetter: () => '',
  });
  try {
    doc.render(data);
  } catch (err) {
    const details = (err.properties && err.properties.errors) || [];
    const detail = details.map((e) => e.properties && e.properties.explanation).filter(Boolean).join('\n');
    throw new Error(detail || err.message || '템플릿 렌더링 중 오류가 발생했습니다.');
  }
  const zip = doc.getZip();
  // .dotx(서식 파일)로 만든 결과는 .docx로 저장되므로 본문 형식도 "문서"로 바꿔야 Word가 연다
  const types = zip.file('[Content_Types].xml');
  if (types && types.asText().includes(DOTX_MAIN_CONTENT_TYPE)) {
    zip.file('[Content_Types].xml', types.asText().split(DOTX_MAIN_CONTENT_TYPE).join(DOCX_MAIN_CONTENT_TYPE));
  }
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/* ---------- 한글 ---------- */

/** 한글 템플릿: 한 번 읽은 모델과 {{변수}} 자리를 두고 행마다 값만 바꿔 쓴다 */
function hwpxTemplate(buffer) {
  const model = modelFromHwpx(buffer);
  const spots = recognizePlaceholders(model);
  return {
    variables: uniqueVariables(spots.map((s) => s.label)),
    // 값은 받은 그대로 넣는다(앞뒤 공백 유지). 없는 변수·제어 태그 자리는 지운다. 줄바꿈은 한글 줄바꿈으로.
    render: (data) =>
      writePlan(
        buffer,
        model,
        spots.map((s) => ({ spotId: s.id, paragraphId: s.paragraphId, start: s.range.start, end: s.range.end, text: Object.prototype.hasOwnProperty.call(data, s.label) ? String(data[s.label]) : '' }))
      ),
  };
}

/* ---------- 템플릿 열기 ---------- */

const templateCache = createFileCache(8);

/**
 * 템플릿 파일을 연다: { path, format, variables, render(data) → Buffer }.
 * 같은 파일은 다시 읽지 않는다(파일이 바뀌면 새로). "변수 추출 → 문서 생성", 일괄 생성의 행마다 쓴다.
 */
function openTemplate(filePath) {
  const format = templateFormat(filePath);
  return templateCache.get(filePath, () => {
    const buffer = safeReadFileSync(filePath);
    if (format === 'hwpx') return { path: filePath, format, ...hwpxTemplate(buffer) };
    return { path: filePath, format, variables: docxVariables(buffer), render: (data) => renderDocx(buffer, data) };
  });
}

/* ---------- 일괄 생성 ---------- */

/** 행 하나의 변수 값: 매핑된 열의 값을 화면에 보이는 글자로(없으면 빈칸) */
function rowData(row, variables, mapping) {
  const data = {};
  variables.forEach((v) => {
    data[v] = formatPreviewValue(mapping[v] ? row[mapping[v]] : '');
  });
  return data;
}

/** Windows/macOS 파일 이름에 쓸 수 없는 글자를 "_"로, 길이는 80자까지 */
function sanitizeFilename(name) {
  let cleaned = formatPreviewValue(name)
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .slice(0, 80)
    .trim()
    // Windows는 끝의 점·공백을 지워 "홍길동."과 "홍길동"이 같은 파일이 된다
    .replace(/[. ]+$/, '');
  // Windows 예약 이름(CON, PRN, AUX, NUL, COM1~9, LPT1~9)은 파일 이름으로 쓸 수 없다
  if (/^(con|prn|aux|nul|com\d|lpt\d)$/i.test(cleaned)) cleaned = `_${cleaned}`;
  return cleaned;
}

/**
 * 행마다 결과 파일 이름. 기준 열 값이 비면 "행N", 같은 이름이 또 나오면 "(2)", "(3)"…
 * (Windows·macOS는 대소문자를 가리지 않으므로 "Kim"과 "kim"도 같은 이름으로 센다).
 * 확인 창 미리보기와 실제 생성이 같은 이름을 쓰도록 이 함수 하나만 쓴다.
 */
function outputFilenames(rows, filenameColumn, extension) {
  const used = new Set();
  return rows.map((row, idx) => {
    const base = (filenameColumn && sanitizeFilename(row[filenameColumn])) || `행${idx + 1}`;
    let name = base;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${base}(${n})`;
    used.add(name.toLowerCase());
    return `${name}.${extension}`;
  });
}

/** 파일 이름 기준 열 값이 비어 행 번호로 이름 짓는 행 수 */
const unnamedRowCount = (rows, filenameColumn) => rows.filter((row) => !filenameColumn || !sanitizeFilename(row[filenameColumn])).length;

/** 변수마다 값이 비어 있는 행 수: { 변수: 개수 } */
function missingValueCounts(rows, variables, mapping) {
  const counts = Object.fromEntries(variables.map((v) => [v, 0]));
  rows.forEach((row) => {
    variables.forEach((v) => {
      const val = mapping[v] ? row[mapping[v]] : '';
      if (val === '' || val === null || val === undefined) counts[v] += 1;
    });
  });
  return counts;
}

/** 데이터 열 중 파일 이름 기준으로 먼저 골라 둘 열("이름"·"성명" 류) */
const defaultFilenameColumn = (headers) => headers.find((h) => /이름|성명|name/i.test(h)) || '';

module.exports = {
  templateFormat,
  isControlToken,
  openTemplate,
  docxVariables,
  renderDocx,
  hwpxTemplate,
  rowData,
  sanitizeFilename,
  outputFilenames,
  unnamedRowCount,
  missingValueCounts,
  defaultFilenameColumn,
};
