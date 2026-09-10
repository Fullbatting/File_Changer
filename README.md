# 문서 변환 및 메일머지 자동화 프로그램

Electron + Node.js + JavaScript로 만든 데스크톱 프로그램입니다. 파이썬은 사용하지 않습니다.

- **탭 1. 문서 양식 변환기**: 엑셀/CSV(`.xlsx`, `.xls`, `.csv`) 파일의 컬럼을 다른 엑셀 양식에 매핑하여 변환
- **탭 2. 메일머지 자동기입**: Word(`.docx`) 또는 한글(`.hwpx`) 템플릿의 `{{변수명}}`을 자동 추출하여 값 입력 후 문서 생성

## 설치 없이 바로 실행하기 (Windows)

1. 이 저장소 페이지에서 **Code → Download ZIP**으로 압축 파일을 내려받습니다.
2. 원하는 폴더에 압축을 풉니다.
3. 압축 푼 폴더 안의 **`run.bat`** 파일을 더블클릭합니다.
   - 최초 실행 시 필요한 패키지를 자동으로 설치합니다 (인터넷 연결 필요, 몇 분 정도 소요).
   - 이후 실행부터는 설치 과정 없이 바로 프로그램이 열립니다.
   - 패키지를 다시 설치하고 싶다면 폴더 안의 `node_modules` 폴더를 삭제한 뒤 `run.bat`을 다시 실행하세요.

> **사전 준비**: PC에 [Node.js](https://nodejs.org/) (LTS 버전)가 설치되어 있어야 합니다. `run.bat` 실행 시 Node.js가 없으면 안내 메시지가 표시됩니다.

## macOS / Linux에서 실행하기

`run.bat`은 Windows 전용입니다. macOS/Linux에서는 터미널에서 프로젝트 폴더로 이동한 뒤 아래 명령을 실행하세요.

```bash
npm install
npm start
```

## 개발자용 명령어

```bash
npm install   # 의존성 설치
npm start     # 앱 실행 (electron .)
```

## 지원 파일 형식

| 기능 | 지원 확장자 |
| --- | --- |
| 문서 양식 변환기 (원본/타겟) | `.xlsx`, `.xls`, `.xlsm`, `.csv` |
| 메일머지 템플릿 | `.docx`, `.dotx`, `.hwpx`, `.hwpt` |

- CSV는 UTF-8(BOM 포함/미포함)과 Windows 한글 기본 인코딩(CP949/EUC-KR)을 자동 판별하여 읽습니다.
- `.hwp`(구형 바이너리 한글 포맷)는 지원하지 않습니다. 한글 프로그램에서 "다른 이름으로 저장 > HWPX"로 변환한 뒤 이용하세요.

## 주요 오픈소스 라이브러리

- [xlsx (SheetJS)](https://github.com/SheetJS/sheetjs) — 엑셀/CSV 처리 (Apache-2.0)
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater) — Word 템플릿 메일머지 (MIT)
- [pizzip](https://github.com/open-xml-templating/pizzip) — zip/OOXML 처리 (MIT)
- [iconv-lite](https://github.com/ashtuchkin/iconv-lite) — CSV 인코딩 변환 (MIT)
