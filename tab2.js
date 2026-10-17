// tab2.js
// TAB 2. 메일머지 자동기입: Word(.docx)/한글(.hwpx) 템플릿 변수 추출 및 치환,
// 직접 입력(1건) / 데이터 파일로 일괄 생성(메일머지) 두 가지 모드

const { ipcRenderer } = require('electron');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const common = require('./common.js');

const {
  log,
  notifyError,
  friendlyErrorMessage,
  safeReadFileSync,
  safeWriteFileSync,
  getFileMtimeMs,
  evictOldestIfFull,
  confirmModal,
  readWorkbookSmart,
  EXCEL_FILTERS,
  DOCX_FILTERS,
  HWPX_FILTERS,
  TEMPLATE_FILTERS,
  bindDropzone,
  refreshSheetSelector,
  extractHeaders,
  findAutoMatch,
  xmlEscape,
  formatPreviewValue: formatMailMergeValue,
} = common;

let templatePath = null;
let templateVars = [];

// 메일머지 입력 방식: 'manual'(1건 직접 입력) | 'batch'(데이터 파일로 일괄 생성)
let mailMergeMode = 'manual';
let mmDataPath = null;
let mmDataSheetName = null;
let mmDataHeaderRowIndex = 0;
let mmDataHeaders = [];
let mmDataRows = [];
let mmColumnMapping = {}; // { 템플릿변수: 데이터컬럼명 }
let mmFilenameColumn = '';

const dropTemplate = document.getElementById('dropTemplate');
const fileNameTemplate = document.getElementById('fileNameTemplate');
const extractStatus = document.getElementById('extractStatus');
const varFormGrid = document.getElementById('varFormGrid');
const manualInputCard = document.getElementById('manualInputCard');
const batchInputCard = document.getElementById('batchInputCard');
const btnGenerateDocxEl = document.getElementById('btnGenerateDocx');
const btnGenerateBatchEl = document.getElementById('btnGenerateBatch');
const generateCardDesc = document.getElementById('generateCardDesc');

bindDropzone(dropTemplate, fileNameTemplate, TEMPLATE_FILTERS, (p) => {
  templatePath = p;
  log(`템플릿 선택됨: ${p}`);
});

document.querySelectorAll('input[name="mmMode"]').forEach((radio) => {
  radio.addEventListener('change', (e) => {
    mailMergeMode = e.target.value;
    const isBatch = mailMergeMode === 'batch';
    manualInputCard.hidden = isBatch;
    batchInputCard.hidden = !isBatch;
    btnGenerateDocxEl.hidden = isBatch;
    btnGenerateBatchEl.hidden = !isBatch;
    generateCardDesc.textContent = isBatch
      ? '데이터 파일의 각 행마다 문서를 생성하여 선택한 폴더에 저장합니다.'
      : '입력한 값으로 템플릿 변수를 치환하여 결과 문서(.docx 또는 .hwpx)를 생성합니다.';
  });
});

/** 파일 확장자로 템플릿 형식을 판별. 지원하지 않는 형식은 안내 메시지와 함께 예외를 던짐 */
function getTemplateFormat(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  // .dotx(Word 서식 파일)는 .docx와, .hwpt(한글 서식 파일)는 .hwpx와 내부 구조가 동일하여 같은 엔진으로 처리
  if (ext === '.docx' || ext === '.dotx') return 'docx';
  if (ext === '.hwpx' || ext === '.hwpt') return 'hwpx';
  if (ext === '.hwp') {
    throw new Error('구형 .hwp(바이너리) 포맷은 지원하지 않습니다. 한글 프로그램에서 "다른 이름으로 저장 > HWPX"로 변환한 뒤 다시 시도하세요.');
  }
  throw new Error('지원하지 않는 템플릿 형식입니다. .docx, .dotx, .hwpx 또는 .hwpt 파일을 선택하세요.');
}

/**
 * docxtemplater는 {{#루프}}...{{/루프}}, {{^조건}}, {{.}} 같은 제어 태그도
 * 같은 {{ }} 구분자를 쓴다. 우리는 반복/조건문 UI를 제공하지 않지만, 혹시
 * 그런 템플릿이 선택되어도 이 태그들이 "변수"로 잘못 추출되어 입력창이나
 * 매핑 테이블에 나타나지 않도록 걸러낸다.
 */
function isTemplateControlToken(varName) {
  return /^[#/^.]/.test(varName.trim());
}

/** docx 파일 내부 word/document.xml 에서 {{변수명}} 패턴을 추출 */
function extractTemplateVariables(filePath) {
  const content = safeReadFileSync(filePath, 'binary');
  const zip = new PizZip(content);

  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('올바른 .docx 파일이 아닙니다. (word/document.xml 없음)');
  }
  const xml = docXmlFile.asText();

  // Word 는 텍스트를 여러 <w:t> 런(run)으로 쪼개는 경우가 많으므로,
  // XML 태그를 제거한 순수 텍스트에서 {{변수명}} 패턴을 탐색한다.
  const plainText = xml.replace(/<[^>]+>/g, '');

  const regex = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const found = new Set();
  let match;
  while ((match = regex.exec(plainText)) !== null) {
    const varName = match[1].trim();
    if (varName && !isTemplateControlToken(varName)) found.add(varName);
  }

  return Array.from(found);
}

/**
 * docx 템플릿 원본 버퍼 캐시. 메일머지 일괄 생성에서 같은 템플릿 파일을
 * 행(row)마다 매번 디스크에서 다시 읽지 않기 위함. docxtemplater는 render()를
 * 호출하면 내부 상태가 치환된 결과로 바뀌어 같은 인스턴스를 재사용할 수
 * 없으므로, 캐시하는 것은 원본 바이트(zip 생성 전)까지만 — PizZip/Docxtemplater
 * 인스턴스는 매번 새로 만들어 행끼리 서로 영향을 주지 않게 한다.
 */
const docxBufferCache = new Map(); // filePath -> { mtimeMs, buffer }
const DOCX_CACHE_LIMIT = 8;

function getDocxTemplateBuffer(filePath) {
  const mtimeMs = getFileMtimeMs(filePath);
  const cached = docxBufferCache.get(filePath);
  if (cached && cached.mtimeMs === mtimeMs) return cached.buffer;

  const buffer = safeReadFileSync(filePath, 'binary');
  evictOldestIfFull(docxBufferCache, DOCX_CACHE_LIMIT);
  docxBufferCache.set(filePath, { mtimeMs, buffer });
  return buffer;
}

/** .docx/.dotx 템플릿의 {{변수명}}을 data 값으로 치환한 결과 파일 버퍼를 생성 */
function renderDocxDocument(filePath, data) {
  const content = getDocxTemplateBuffer(filePath);
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
  });

  try {
    doc.render(data);
  } catch (renderErr) {
    const details = (renderErr.properties && renderErr.properties.errors) || [];
    const detailMsg = details.map((e) => e.properties && e.properties.explanation).filter(Boolean).join('\n');
    throw new Error(detailMsg || renderErr.message || '템플릿 렌더링 중 오류가 발생했습니다.');
  }

  return doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/** 템플릿 형식에 맞는 렌더링 함수를 호출하는 공용 디스패처 (직접 입력/일괄 생성 모두 사용) */
function renderMergedDocument(filePath, format, data) {
  return format === 'hwpx' ? renderHwpxDocument(filePath, data) : renderDocxDocument(filePath, data);
}

/* -------------------------------------------------------------------------
 * 한글(HWPX) 템플릿 처리
 *
 * .hwpx는 .docx와 마찬가지로 zip 컨테이너 + XML 구조를 사용한다. 본문은
 * Contents/section0.xml, section1.xml, ... 에 문단(<hp:p>) > 런(<hp:run>) >
 * 텍스트(<hp:t>) 계층으로 저장되며, 워드프로세서가 입력 도중 하나의 문구를
 * 여러 <hp:t> 런으로 쪼개는 경우가 있어 .docx와 동일하게 "모든 런의 텍스트를
 * 이어붙여 탐색 → 매치된 구간을 원래 런 위치에 되돌려 치환" 방식으로 처리한다.
 * (docx에는 docxtemplater라는 성숙한 라이브러리가 있지만, HWPX를 지원하는
 * 오픈소스 JS 라이브러리는 없어 동일한 원리를 직접 구현했다.)
 * ------------------------------------------------------------------------- */

/** hwpx zip에서 본문 섹션 XML 파일 이름 목록을 번호순으로 반환 */
function getHwpxSectionNames(zip) {
  return Object.keys(zip.files)
    .filter((name) => /^Contents\/section\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/section(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/section(\d+)\.xml/)[1], 10);
      return na - nb;
    });
}

/** section XML 하나에서 <hp:t> 런들의 원본 위치와, 이를 이어붙인 순수 텍스트를 함께 추출 */
function parseHwpxRuns(xml) {
  const runRegex = /<hp:t\b([^>]*)>([\s\S]*?)<\/hp:t>/g;
  const runs = [];
  let match;
  let offset = 0;
  while ((match = runRegex.exec(xml)) !== null) {
    const content = match[2];
    runs.push({
      xmlStart: match.index,
      xmlEnd: match.index + match[0].length,
      openTag: `<hp:t${match[1]}>`,
      content,
      plainStart: offset,
      plainEnd: offset + content.length,
    });
    offset += content.length;
  }
  return { runs, plainText: runs.map((r) => r.content).join('') };
}

/** 이어붙인 순수 텍스트에서 {{변수명}} 목록을 추출 */
function extractVariablesFromPlainText(plainText) {
  const regex = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const found = [];
  let match;
  while ((match = regex.exec(plainText)) !== null) {
    const varName = match[1].trim();
    if (varName && !isTemplateControlToken(varName)) found.push(varName);
  }
  return found;
}

/**
 * section XML 안의 {{변수명}}을 data 값으로 치환한 새 XML 문자열을 반환.
 * runs/plainText는 parseHwpxRuns(xml)의 결과를 그대로 전달받는다 — 변수 추출
 * 단계에서 이미 계산해 둔 값을 재사용해 같은 정규식 스캔을 두 번 하지 않기 위함.
 */
function renderHwpxXml(xml, runs, plainText, data) {
  if (runs.length === 0) return xml;

  const regex = /\{\{\s*([^{}]+?)\s*\}\}/g;
  const runEdits = runs.map(() => []); // [{ localStart, localEnd, replacement }]
  let match;

  while ((match = regex.exec(plainText)) !== null) {
    const varName = match[1].trim();
    const matchStart = match.index;
    const matchEnd = matchStart + match[0].length;
    const value = xmlEscape(Object.prototype.hasOwnProperty.call(data, varName) ? data[varName] : '');

    let firstTouchedRun = true;
    runs.forEach((run, i) => {
      const overlapStart = Math.max(matchStart, run.plainStart);
      const overlapEnd = Math.min(matchEnd, run.plainEnd);
      if (overlapStart < overlapEnd) {
        // 치환 값은 매치가 처음 걸친 런에만 삽입하고, 나머지 겹치는 런에서는 해당 구간만 비운다.
        runEdits[i].push({
          localStart: overlapStart - run.plainStart,
          localEnd: overlapEnd - run.plainStart,
          replacement: firstTouchedRun ? value : '',
        });
        firstTouchedRun = false;
      }
    });
  }

  let result = '';
  let cursor = 0;
  runs.forEach((run, i) => {
    result += xml.slice(cursor, run.xmlStart);
    let content = run.content;
    // 뒤쪽 구간부터 잘라내야 앞쪽 구간의 인덱스가 밀리지 않는다.
    runEdits[i]
      .sort((a, b) => b.localStart - a.localStart)
      .forEach((e) => {
        content = content.slice(0, e.localStart) + e.replacement + content.slice(e.localEnd);
      });
    result += run.openTag + content + '</hp:t>';
    cursor = run.xmlEnd;
  });
  result += xml.slice(cursor);
  return result;
}

/**
 * hwpx 파싱 캐시. "변수 자동 추출" 다음 "문서 생성"으로 이어지는 흐름에서
 * 같은 템플릿 파일을 매번 디스크에서 다시 읽고, zip을 다시 열고, 각 섹션의
 * <hp:t> 런을 다시 정규식으로 스캔하는 중복을 없애기 위함. 파일의 mtime이
 * 바뀌면 자동으로 무효화된다. 캐시된 섹션의 xml/runs/plainText는 원본
 * 그대로(치환 전) 보존되므로, 값을 바꿔 여러 번 "문서 생성"을 눌러도 매번
 * 정확히 원본 플레이스홀더 기준으로 다시 치환된다 — zip 객체 자체는 재사용하되
 * 매 생성마다 섹션 내용을 원본 기준으로 새로 덮어쓰기 때문에 안전하다.
 */
const hwpxCache = new Map(); // filePath -> { mtimeMs, zip, sections: [{ name, xml, runs, plainText }] }
const HWPX_CACHE_LIMIT = 8;

function getHwpxCacheEntry(filePath) {
  const mtimeMs = getFileMtimeMs(filePath);
  const cached = hwpxCache.get(filePath);
  if (cached && cached.mtimeMs === mtimeMs) {
    return cached;
  }

  const content = safeReadFileSync(filePath);
  const zip = new PizZip(content);
  const sectionNames = getHwpxSectionNames(zip);
  if (sectionNames.length === 0) {
    throw new Error('올바른 .hwpx 파일이 아닙니다. (Contents/section*.xml 없음)');
  }
  const sections = sectionNames.map((name) => {
    const xml = zip.file(name).asText();
    const { runs, plainText } = parseHwpxRuns(xml);
    return { name, xml, runs, plainText };
  });

  const entry = { mtimeMs, zip, sections };
  evictOldestIfFull(hwpxCache, HWPX_CACHE_LIMIT);
  hwpxCache.set(filePath, entry);
  return entry;
}

/** .hwpx 템플릿에서 {{변수명}} 목록을 추출 */
function extractHwpxVariables(filePath) {
  const entry = getHwpxCacheEntry(filePath);
  const found = new Set();
  entry.sections.forEach(({ plainText }) => {
    extractVariablesFromPlainText(plainText).forEach((v) => found.add(v));
  });
  return Array.from(found);
}

/** .hwpx 템플릿의 {{변수명}}을 data 값으로 치환한 결과 파일 버퍼를 생성 */
function renderHwpxDocument(filePath, data) {
  const entry = getHwpxCacheEntry(filePath);
  entry.sections.forEach(({ name, xml, runs, plainText }) => {
    entry.zip.file(name, renderHwpxXml(xml, runs, plainText, data));
  });
  return entry.zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const varInputs = {}; // { varName: <input> }

function renderVarForm() {
  varFormGrid.innerHTML = '';
  Object.keys(varInputs).forEach((k) => delete varInputs[k]);

  if (templateVars.length === 0) {
    varFormGrid.innerHTML = '<div class="empty-hint">템플릿에서 {{변수명}} 형식의 항목을 찾지 못했습니다.</div>';
    return;
  }

  templateVars.forEach((varName) => {
    const wrap = document.createElement('div');
    wrap.className = 'var-field';

    const label = document.createElement('label');
    label.innerHTML = `<span class="tag">{{${varName}}}</span>`;
    wrap.appendChild(label);

    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = `${varName} 값을 입력하세요`;
    input.dataset.varName = varName;
    wrap.appendChild(input);

    varFormGrid.appendChild(wrap);
    varInputs[varName] = input;
  });
}

/* -------------------------------------------------------------------------
 * 메일머지 일괄 생성 (데이터 파일의 각 행마다 문서 하나씩 생성)
 * 엑셀/CSV 읽기·시트 선택·헤더 행 자동 감지·편집거리 기반 컬럼 자동 매핑은
 * 공통 모듈(common.js)의 readWorkbookSmart/extractHeaders/refreshSheetSelector/
 * findAutoMatch를 그대로 재사용한다 (탭1의 컬럼 매핑과 동일한 로직).
 * ------------------------------------------------------------------------- */

const dropMailMergeData = document.getElementById('dropMailMergeData');
const fileNameMailMergeData = document.getElementById('fileNameMailMergeData');
const sheetRowMailMerge = document.getElementById('sheetRowMailMerge');
const sheetSelectMailMerge = document.getElementById('sheetSelectMailMerge');
const mailMergeDataStatus = document.getElementById('mailMergeDataStatus');
const mailMergeMappingTableBody = document.getElementById('mailMergeMappingTableBody');
const filenameColumnSelect = document.getElementById('filenameColumnSelect');

sheetRowMailMerge.addEventListener('click', (e) => e.stopPropagation());

/** Windows/macOS 파일명에 쓸 수 없는 문자를 "_"로 바꾸고 길이를 적절히 제한 */
function sanitizeFilename(name) {
  const cleaned = String(name).trim().replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_');
  return cleaned.slice(0, 80);
}

/**
 * 각 행의 결과 파일명을 미리 계산한다. filenameColumn 값이 비어 있으면
 * "행N"으로 대체하고, 같은 이름이 여러 번 나오면 "(2)", "(3)"...을 붙여
 * 서로 다른 파일로 구분한다. 검증 미리보기와 실제 생성에서 동일한 결과가
 * 나오도록 이 함수 하나만 사용한다.
 */
function computeOutputFilenames(rows, filenameColumn, extension) {
  const usedCount = new Map();
  return rows.map((row, idx) => {
    let base = filenameColumn ? sanitizeFilename(row[filenameColumn]) : '';
    if (!base) base = `행${idx + 1}`;
    const count = usedCount.get(base) || 0;
    usedCount.set(base, count + 1);
    const finalName = count === 0 ? base : `${base}(${count + 1})`;
    return `${finalName}.${extension}`;
  });
}

/** 데이터 파일/시트가 (재)선택되면 이전에 불러온 행·매핑이 새 시트와 어긋나지 않도록 초기화 */
function resetMailMergeDataState() {
  mmDataHeaders = [];
  mmDataRows = [];
  renderMailMergeMappingTable();
  mailMergeDataStatus.textContent = '대기 중';
  mailMergeDataStatus.className = 'pill warn';
}

bindDropzone(dropMailMergeData, fileNameMailMergeData, EXCEL_FILTERS, (p) => {
  mmDataPath = p;
  log(`메일머지 데이터 파일 선택됨: ${p}`);
  resetMailMergeDataState();
  refreshSheetSelector(
    p,
    sheetRowMailMerge,
    sheetSelectMailMerge,
    (name) => {
      mmDataSheetName = name;
    },
    (name) => {
      resetMailMergeDataState();
      log(`시트가 "${name}"(으)로 변경되었습니다. [데이터 불러오기]를 다시 실행하세요.`, 'warn');
    }
  );
});

function renderFilenameColumnOptions() {
  filenameColumnSelect.innerHTML = '<option value="">-- 행 번호 사용 --</option>';
  mmDataHeaders.forEach((col) => {
    const opt = document.createElement('option');
    opt.value = col;
    opt.textContent = col;
    filenameColumnSelect.appendChild(opt);
  });
  // "이름"/"성명" 류 컬럼이 있으면 파일명 기준으로 기본 선택 (편의 기능)
  const nameLike = mmDataHeaders.find((h) => /이름|성명|name/i.test(h));
  mmFilenameColumn = nameLike || '';
  filenameColumnSelect.value = mmFilenameColumn;
  filenameColumnSelect.onchange = () => {
    mmFilenameColumn = filenameColumnSelect.value;
  };
}

const MAPPING_PREVIEW_SAMPLE_SIZE = 2;

/** mmDataRows의 앞부분 몇 행에서 해당 컬럼의 실제 값을 뽑아 미리보기 텍스트로 만든다 */
function buildMailMergeColumnPreview(column) {
  if (!column) return '';
  return mmDataRows
    .slice(0, MAPPING_PREVIEW_SAMPLE_SIZE)
    .map((row) => formatMailMergeValue(row[column]))
    .filter((v) => v !== '')
    .join(' / ');
}

/**
 * 템플릿 변수 ← 데이터 컬럼 매핑 테이블을 그린다 (공통 findAutoMatch로 자동 매핑 제안).
 * 매핑된 데이터 컬럼의 실제 값(앞 2행)을 미리보기로 함께 보여줘, 자동/수동 매핑이
 * 맞는지 일괄 생성을 누르기 전에 바로 확인할 수 있게 한다 (CSV/엑셀 임포트 도구들의
 * "Match Columns" 단계에서 흔히 쓰이는 패턴).
 */
function renderMailMergeMappingTable() {
  mailMergeMappingTableBody.innerHTML = '';
  mmColumnMapping = {};

  if (templateVars.length === 0 || mmDataHeaders.length === 0) {
    mailMergeMappingTableBody.innerHTML =
      '<tr><td colspan="3" class="empty-hint">먼저 [변수 자동 추출]과 [데이터 불러오기]를 실행하세요.</td></tr>';
    return;
  }

  const usedColumns = new Set();
  templateVars.forEach((varName) => {
    const tr = document.createElement('tr');

    const tdVar = document.createElement('td');
    tdVar.textContent = `{{${varName}}}`;
    tr.appendChild(tdVar);

    const tdSelect = document.createElement('td');
    const select = document.createElement('select');
    select.dataset.templateVar = varName;

    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = '-- 선택 안함 (빈 값) --';
    select.appendChild(emptyOpt);

    mmDataHeaders.forEach((col) => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      select.appendChild(opt);
    });

    const tdPreview = document.createElement('td');
    tdPreview.className = 'mapping-preview-cell';

    const { match: autoMatch, reason } = findAutoMatch(varName, mmDataHeaders, usedColumns);
    if (autoMatch) {
      select.value = autoMatch;
      mmColumnMapping[varName] = autoMatch;
      usedColumns.add(autoMatch);
      if (reason === 'synonym') {
        log(`변수 "${varName}"을 동의어 사전에 따라 데이터 컬럼 "${autoMatch}"에 자동 매핑했습니다.`);
      } else if (reason === 'fuzzy') {
        log(`변수 "${varName}"을 이름이 비슷한 데이터 컬럼 "${autoMatch}"에 자동 매핑했습니다. 맞는지 확인하세요.`, 'warn');
      }
    } else {
      mmColumnMapping[varName] = '';
    }
    tdPreview.textContent = buildMailMergeColumnPreview(mmColumnMapping[varName]);

    select.addEventListener('change', () => {
      mmColumnMapping[varName] = select.value;
      tdPreview.textContent = buildMailMergeColumnPreview(select.value);
    });

    tdSelect.appendChild(select);
    tr.appendChild(tdSelect);
    tr.appendChild(tdPreview);
    mailMergeMappingTableBody.appendChild(tr);
  });
}

document.getElementById('btnLoadMailMergeData').addEventListener('click', () => {
  try {
    if (!mmDataPath) throw new Error('데이터 파일을 먼저 선택하세요.');
    if (templateVars.length === 0) throw new Error('먼저 템플릿을 선택하고 [변수 자동 추출]을 실행하세요.');

    log(`메일머지 데이터를 불러옵니다... (시트: ${mmDataSheetName || '첫 번째 시트'})`);

    const result = extractHeaders(mmDataPath, mmDataSheetName);
    mmDataHeaders = result.headers;
    mmDataHeaderRowIndex = result.headerRowIndex;
    if (mmDataHeaderRowIndex > 0) {
      log(`제목/빈 행으로 보이는 ${mmDataHeaderRowIndex}개 행을 건너뛰고 ${mmDataHeaderRowIndex + 1}행을 헤더로 인식했습니다.`, 'warn');
    }

    const workbook = readWorkbookSmart(mmDataPath);
    const sheet = workbook.Sheets[mmDataSheetName || workbook.SheetNames[0]];
    mmDataRows = XLSX.utils.sheet_to_json(sheet, { defval: '', range: mmDataHeaderRowIndex });
    if (mmDataRows.length === 0) throw new Error('데이터 파일에 변환할 행이 없습니다.');

    renderMailMergeMappingTable();
    renderFilenameColumnOptions();

    mailMergeDataStatus.textContent = `${mmDataRows.length}행 불러옴`;
    mailMergeDataStatus.className = 'pill ok';
    log(`데이터 ${mmDataRows.length}행, 컬럼 ${mmDataHeaders.length}개를 불러왔습니다.`, 'ok');
  } catch (err) {
    mailMergeDataStatus.textContent = '불러오기 실패';
    mailMergeDataStatus.className = 'pill warn';
    notifyError('메일머지 데이터를 불러오는 중 오류가 발생했습니다', err);
  }
});

/** 템플릿 변수별로 값이 비어 있는 행 수를 집계 (저장 전 검증) */
function validateBatchRows(rows, templateVarsList, columnMapping) {
  const missingByVar = {};
  templateVarsList.forEach((v) => {
    missingByVar[v] = 0;
  });
  rows.forEach((row) => {
    templateVarsList.forEach((v) => {
      const col = columnMapping[v];
      const val = col ? row[col] : '';
      if (val === '' || val === null || val === undefined) missingByVar[v] += 1;
    });
  });
  return { missingByVar };
}

// 메일머지 일괄 생성 확인(검증 + 미리보기) 모달
const mailMergeReviewModal = document.getElementById('mailMergeReviewModal');
const mailMergeReviewSummary = document.getElementById('mailMergeReviewSummary');
const mailMergeReviewTableHead = document.getElementById('mailMergeReviewTableHead');
const mailMergeReviewTableBody = document.getElementById('mailMergeReviewTableBody');
const mailMergeReviewProceedBtn = document.getElementById('mailMergeReviewProceed');
const mailMergeReviewCancelBtn = document.getElementById('mailMergeReviewCancel');

const MAILMERGE_PREVIEW_ROW_LIMIT = 5;

/** 일괄 생성 확인 모달을 띄운다. "그대로 생성"(true) / "취소하고 수정"(false)을 Promise로 반환 */
function showMailMergeReviewModal(templateVarsList, rows, columnMapping, filenames, validation) {
  const lines = [];
  lines.push(`<div class="review-line ok">데이터 ${rows.length}행 → 문서 <b>${rows.length}</b>개를 생성합니다.</div>`);

  const missingEntries = Object.entries(validation.missingByVar).filter(([, count]) => count > 0);
  if (missingEntries.length > 0) {
    const detail = missingEntries.map(([v, count]) => `"${xmlEscape(v)}" ${count}건`).join(', ');
    lines.push(`<div class="review-line warn">⚠ 값이 비어 있는 변수: ${detail}</div>`);
  } else {
    lines.push(`<div class="review-line ok">✓ 모든 변수에 값이 채워져 있습니다.</div>`);
  }

  const fallbackCount = rows.filter((row) => !mmFilenameColumn || !String(row[mmFilenameColumn] || '').trim()).length;
  if (fallbackCount > 0) {
    lines.push(`<div class="review-line warn">⚠ 파일명 기준 컬럼 값이 비어 있어 행 번호로 이름 지어지는 행: ${fallbackCount}건</div>`);
  }

  mailMergeReviewSummary.innerHTML = lines.join('');

  mailMergeReviewTableHead.innerHTML = ['파일명', ...templateVarsList].map((h) => `<th>${xmlEscape(h)}</th>`).join('');

  const previewCount = Math.min(MAILMERGE_PREVIEW_ROW_LIMIT, rows.length);
  const previewHtml = [];
  for (let idx = 0; idx < previewCount; idx++) {
    const row = rows[idx];
    const cells = [
      filenames[idx],
      ...templateVarsList.map((v) => {
        const col = columnMapping[v];
        const val = col ? row[col] : '';
        return formatMailMergeValue(val);
      }),
    ];
    previewHtml.push(`<tr>${cells.map((c) => `<td>${xmlEscape(c)}</td>`).join('')}</tr>`);
  }
  if (rows.length > previewCount) {
    previewHtml.push(`<tr><td colspan="${templateVarsList.length + 1}" class="empty-hint">... 외 ${rows.length - previewCount}행</td></tr>`);
  }
  mailMergeReviewTableBody.innerHTML = previewHtml.join('');

  return confirmModal(mailMergeReviewModal, mailMergeReviewProceedBtn, mailMergeReviewCancelBtn);
}

btnGenerateBatchEl.addEventListener('click', async () => {
  try {
    if (!templatePath) throw new Error('Word/한글(.docx, .hwpx) 템플릿 파일을 먼저 선택하세요.');
    if (templateVars.length === 0) throw new Error('추출된 변수가 없습니다. 먼저 [변수 자동 추출]을 실행하세요.');
    if (mmDataRows.length === 0) throw new Error('데이터 파일을 먼저 불러오세요.');

    const format = getTemplateFormat(templatePath);
    const mappedCount = templateVars.filter((v) => mmColumnMapping[v]).length;
    if (mappedCount === 0) throw new Error('매핑된 변수가 하나도 없습니다. 컬럼 매핑을 확인하세요.');

    const validation = validateBatchRows(mmDataRows, templateVars, mmColumnMapping);
    const filenames = computeOutputFilenames(mmDataRows, mmFilenameColumn, format);

    const proceed = await showMailMergeReviewModal(templateVars, mmDataRows, mmColumnMapping, filenames, validation);
    if (!proceed) {
      log('사용자가 확인 단계에서 취소했습니다. 매핑을 수정한 뒤 다시 시도하세요.', 'warn');
      return;
    }

    const outputDir = await ipcRenderer.invoke('dialog:selectFolder', { title: '결과 파일을 저장할 폴더 선택' });
    if (!outputDir) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }

    log(`일괄 생성을 시작합니다... (${mmDataRows.length}건, 형식: ${format.toUpperCase()})`);

    // 수백~수천 행을 한 번에 생성할 때 렌더러(UI) 스레드가 완전히 멈춰 보이지
    // 않도록, 일정 행마다 이벤트 루프에 제어를 한 번씩 양보한다.
    const YIELD_EVERY = 20;
    const PROGRESS_LOG_EVERY = 100;

    let successCount = 0;
    const errors = [];
    for (let idx = 0; idx < mmDataRows.length; idx++) {
      const row = mmDataRows[idx];
      try {
        const data = {};
        templateVars.forEach((v) => {
          const col = mmColumnMapping[v];
          const val = col ? row[col] : '';
          data[v] = formatMailMergeValue(val);
        });
        const outputBuffer = renderMergedDocument(templatePath, format, data);
        safeWriteFileSync(path.join(outputDir, filenames[idx]), outputBuffer);
        successCount += 1;
      } catch (err) {
        errors.push({ row: idx + 1, message: friendlyErrorMessage(err) });
      }

      const done = idx + 1;
      if (done % PROGRESS_LOG_EVERY === 0 && done < mmDataRows.length) {
        log(`진행 중... (${done}/${mmDataRows.length})`);
      }
      if (done % YIELD_EVERY === 0) {
        await new Promise((resolve) => setTimeout(resolve, 0));
      }
    }

    const resultLevel = errors.length > 0 ? 'warn' : 'ok';
    log(
      `일괄 생성 완료: 총 ${mmDataRows.length}건 중 ${successCount}건 성공${errors.length > 0 ? `, ${errors.length}건 실패` : ''}.`,
      resultLevel
    );
    errors.slice(0, 5).forEach((e) => log(`  - ${e.row}행 실패: ${e.message}`, 'warn'));
    if (errors.length > 5) log(`  ... 외 ${errors.length - 5}건 실패`, 'warn');

    if (errors.length === 0) {
      alert(`✅ 총 ${successCount}개의 문서를 생성했습니다.\n저장 위치: ${outputDir}`);
    } else {
      alert(`⚠️ ${successCount}개 생성 완료, ${errors.length}건 실패했습니다.\n자세한 내용은 로그를 확인하세요.\n저장 위치: ${outputDir}`);
    }
  } catch (err) {
    notifyError('일괄 생성 중 오류가 발생했습니다', err);
  }
});

document.getElementById('btnExtractVars').addEventListener('click', () => {
  try {
    if (!templatePath) {
      throw new Error('Word/한글(.docx, .hwpx) 템플릿 파일을 먼저 선택하세요.');
    }
    const format = getTemplateFormat(templatePath);
    log(`템플릿 변수 추출을 시작합니다... (형식: ${format.toUpperCase()})`);
    templateVars = format === 'hwpx' ? extractHwpxVariables(templatePath) : extractTemplateVariables(templatePath);

    if (templateVars.length === 0) {
      extractStatus.textContent = '변수 없음';
      extractStatus.className = 'pill warn';
      log('템플릿에서 {{변수명}} 패턴을 찾지 못했습니다.', 'warn');
    } else {
      extractStatus.textContent = `${templateVars.length}개 변수 발견`;
      extractStatus.className = 'pill ok';
      log(`변수 ${templateVars.length}개를 발견했습니다: ${templateVars.join(', ')}`, 'ok');
    }

    renderVarForm();
    renderMailMergeMappingTable();
  } catch (err) {
    extractStatus.textContent = '추출 실패';
    extractStatus.className = 'pill warn';
    notifyError('변수 추출 중 오류가 발생했습니다', err);
  }
});

btnGenerateDocxEl.addEventListener('click', async () => {
  try {
    if (!templatePath) {
      throw new Error('Word/한글(.docx, .hwpx) 템플릿 파일을 먼저 선택하세요.');
    }
    if (templateVars.length === 0) {
      throw new Error('추출된 변수가 없습니다. 먼저 [변수 자동 추출]을 실행하세요.');
    }
    const format = getTemplateFormat(templatePath);

    const data = {};
    const emptyVars = [];
    templateVars.forEach((varName) => {
      const input = varInputs[varName];
      data[varName] = input ? input.value : '';
      if (!data[varName]) emptyVars.push(varName);
    });
    if (emptyVars.length > 0) {
      log(`값이 입력되지 않은 변수 ${emptyVars.length}개는 빈 문자열로 치환됩니다: ${emptyVars.join(', ')}`, 'warn');
    }

    log(`메일머지 문서 생성을 시작합니다... (형식: ${format.toUpperCase()})`);

    const outputBuffer = renderMergedDocument(templatePath, format, data);

    const baseName = path.basename(templatePath, path.extname(templatePath));
    const savePath = await ipcRenderer.invoke('dialog:saveFile', {
      title: '메일머지 결과 저장',
      defaultPath: `${baseName}_결과.${format}`,
      filters: format === 'hwpx' ? HWPX_FILTERS : DOCX_FILTERS,
    });
    if (!savePath) {
      log('저장이 취소되었습니다.', 'warn');
      return;
    }

    safeWriteFileSync(savePath, outputBuffer);
    log(`메일머지 문서가 저장되었습니다: ${savePath}`, 'ok');
    alert('✅ 문서 생성이 완료되었습니다.');
  } catch (err) {
    notifyError('문서 생성 중 오류가 발생했습니다', err);
  }
});
