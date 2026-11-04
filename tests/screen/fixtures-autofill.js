// 서식 자동 기입 시나리오용 서식: 실제 메일머지 서식처럼 파일 이름을 짓고(기관-서식(세대주).hwpx), 본문에는
// 한글이 .hwp를 .hwpx로 바꿀 때처럼 메일머지 필드(fieldBegin … {{이름}} … fieldEnd)를 넣는다. 글자는 모두 가짜 예시다.
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

const NS = 'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"';
let fieldId = 100;
const F = (name) => {
  fieldId += 1;
  return `<hp:run charPrIDRef="0"><hp:ctrl><hp:fieldBegin id="${fieldId}" type="MAILMERGE" name="${name}" editable="1" dirty="0"/></hp:ctrl><hp:t>{{${name}}}</hp:t><hp:ctrl><hp:fieldEnd beginIDRef="${fieldId}"/></hp:ctrl></hp:run>`;
};
const T = (t) => `<hp:run charPrIDRef="0"><hp:t>${t}</hp:t></hp:run>`;
const P = (...runs) => `<hp:p id="0" paraPrIDRef="0" styleIDRef="0">${runs.join('')}</hp:p>`;
const digits = (order) => P(...order.map((n) => F(String(n))));

function write(dir, name, paragraphs) {
  const zip = new PizZip();
  zip.file('mimetype', 'application/hwp+zip', { compression: 'STORE' });
  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec ${NS}>${paragraphs.join('')}</hs:sec>`);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, name), zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

const root = path.join(process.env.SCREEN_FIX, 'forms');
const emergency = (withHead) => [
  P(T('[서식 1호] ')),
  P(T('현 장 확 인 서')),
  P(T('성 명 '), F('이름'), T(' 생년월일 '), F('생년월일')),
  P(T('주 소 '), F('주소'), T(' (연락처:'), F('연락처'), T(')')),
  P(T('2026년 '), F('월1'), F('월2'), T('월 '), F('일1'), F('일2'), T('일')),
  P(T('[서식 3호]')),
  P(T('금융정보 등(금융･신용･보험정보) 제공 동의서')),
  ...(withHead ? [digits([4, 5, 6, 1, 2, 3, 7, 8, 9, 10, 11, 12, 13])] : []),
  P(T('관계 본인 성명 '), F('이름'), T(' 주민등록번호 '), F('주민등록번호')),
  P(T('[별지 4]')),
  P(T('진료비 확인 요청 위임장(보장기관)')),
  P(T('의료기관명 [예시 병원] 진료기간 '), F('진료기간')),
];
write(path.join(root, '1. (세대주)의료비지원 신청서-메일머지'), '양주시청-긴급복지지원 신청서(세대주).hwpx', emergency(true));
write(path.join(root, '2. (세대원)의료비지원 신청서-메일머지'), '양주시청-긴급복지지원 신청서(세대원).hwpx', emergency(false));
write(path.join(root, '보건소 암지원 신청서'), '공통_암환자의료비등록신청서(메일머지).hwpx', [
  P(T('[별지 제1호서식]')),
  P(T('암환자 의료비 등록 신청서')),
  P(T('성명 '), F('이름'), T(' 주민등록번호 '), F('주민등록번호')),
  P(T('주질병 '), F('주질병'), T(' 질병분류기호 '), F('질병분류기호'), T(' 진단일 '), F('진단일')),
  P(T('담당자 '), F('담당자'), T(' ('), F('담당자 번호'), T(')')),
  P(T('보호자 '), F('보호자')),
]);
console.log('wrote autofill forms');
