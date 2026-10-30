const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const xml = '<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">' +
  '<hp:p><hp:run charPrIDRef="0"><hp:t>안녕하세요 {{이</hp:t></hp:run><hp:run charPrIDRef="1"><hp:t>름}}님</hp:t></hp:run></hp:p></hs:sec>';
const zip = new PizZip();
zip.file('mimetype', 'application/hwp+zip');
zip.file('Contents/section0.xml', xml);
fs.writeFileSync(path.join(process.env.LEGACY_FIX, 'template.hwpx'), zip.generate({ type: 'nodebuffer' }));
