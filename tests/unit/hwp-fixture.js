// tests/unit/hwp-fixture.js
// 테스트용 .hwp(HWP 5.0) 만들기: OLE 복합 문서에 FileHeader와 BodyText/Section0을 넣고, 본문은 레코드로 쓴다.
// 문단 하나 = PARA_HEADER + PARA_TEXT(+ 메일머지 필드면 CTRL_HEADER '%mmg'). 실제 한글 파일처럼 필드 글자는
// 필드 시작(0x03)·끝(0x04) 제어 문자 사이에 "{{이름}}"으로 둔다. 글자는 모두 가짜 예시다.

const zlib = require('zlib');
const XLSX = require('xlsx');

const TAG = { PARA_HEADER: 66, PARA_TEXT: 67, CTRL_HEADER: 71 };
const MAILMERGE = 0x256d6d67; // '%mmg'

function record(tag, level, payload) {
  const size = payload.length;
  const big = size >= 0xfff;
  const head = Buffer.alloc(big ? 8 : 4);
  head.writeUInt32LE((tag & 0x3ff) | ((level & 0x3ff) << 10) | ((big ? 0xfff : size) << 20), 0);
  if (big) head.writeUInt32LE(size, 4);
  return Buffer.concat([head, payload]);
}

/** 8글자(16바이트) 크기 제어 문자: 코드, 컨트롤 ID, 나머지 0, 끝에 같은 코드 */
function wideControl(code, ctrlId) {
  const b = Buffer.alloc(16);
  b.writeUInt16LE(code, 0);
  b.writeUInt32LE(ctrlId, 2);
  b.writeUInt16LE(code, 14);
  return b;
}

const utf16 = (s) => Buffer.from(s, 'utf16le');

/**
 * 문단 하나. parts: 글자(string) 또는 { field: '이름' }(메일머지 필드).
 * 반환: 레코드 버퍼들(문단 머리 → 글자 → 필드 컨트롤)
 */
function paragraph(parts) {
  const text = [];
  const ctrls = [];
  [].concat(parts).forEach((p) => {
    if (typeof p === 'string') {
      text.push(utf16(p));
      return;
    }
    text.push(wideControl(3, MAILMERGE), utf16(`{{${p.field}}}`), wideControl(4, MAILMERGE));
    const name = utf16(`${p.field}${p.extra ? `\r${p.extra}` : ''}`);
    const head = Buffer.alloc(11);
    head.writeUInt32LE(MAILMERGE, 0);
    head.writeUInt16LE(name.length / 2, 9);
    ctrls.push(record(TAG.CTRL_HEADER, 1, Buffer.concat([head, name])));
  });
  text.push(Buffer.from([13, 0])); // 문단 끝
  return [record(TAG.PARA_HEADER, 0, Buffer.alloc(22)), record(TAG.PARA_TEXT, 1, Buffer.concat(text)), ...ctrls];
}

/** paragraphs: paragraph()에 넘길 parts 목록들. compressed: 본문을 압축(실제 한글 기본) */
function buildHwp(paragraphs, { compressed = true } = {}) {
  const header = Buffer.alloc(256);
  header.write('HWP Document File', 0, 'latin1');
  header[35] = 5; // 5.0
  header.writeUInt32LE(compressed ? 1 : 0, 36);
  let body = Buffer.concat(paragraphs.flatMap(paragraph));
  if (compressed) body = zlib.deflateRawSync(body);
  const cfb = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(cfb, '/FileHeader', header);
  XLSX.CFB.utils.cfb_add(cfb, '/BodyText/Section0', body);
  return Buffer.from(XLSX.CFB.write(cfb, { type: 'buffer', fileType: 'cfb' }));
}

module.exports = { buildHwp };
