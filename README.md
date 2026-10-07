# 문서 변환 및 메일머지 자동화 프로그램

Electron + Node.js + JavaScript로 만든 데스크톱 프로그램입니다. 파이썬은 사용하지 않습니다.

- **탭 1. 문서 양식 변환기**: 엑셀/CSV(`.xlsx`, `.xls`, `.csv`) 파일의 컬럼을 다른 엑셀 양식에 매핑하여 변환. 매핑 규칙은 이름 붙여 프리셋으로 저장해 다음에 바로 불러오거나 수정할 수 있습니다. 내부적으로 Extract(읽기) → Transform(매핑) → Load(서식 적용 저장)의 ETL 구조로 동작하며, 결과 파일은 헤더 색상/테두리와 자동 맞춤 열 너비, 숫자 열 천단위 서식이 적용된 상태로 생성됩니다.
  - 컬럼 자동 매핑: 이름이 똑같지 않아도 편집거리가 가까운 컬럼(표기 차이·오타 수준)을 자동으로 찾아 매핑을 제안합니다. 모호하면(후보가 여러 개 동률) 자동 매핑하지 않습니다.
  - 헤더 행 자동 감지: 1행에 "2026년 명단"류 제목이 있어 실제 헤더가 그 아래 행에 있어도 자동으로 찾아냅니다.
  - 저장 전 검증: B 컬럼을 "필수"로 지정하면 해당 값이 빈 행을 저장 전에 알려주고, 모든 값이 완전히 같은 중복 행도 감지합니다. 결과 미리보기를 함께 확인한 뒤 저장을 진행하거나 취소해 매핑을 수정할 수 있습니다.
- **탭 2. 메일머지 자동기입**: Word(`.docx`) 또는 한글(`.hwpx`) 템플릿의 `{{변수명}}`을 자동 추출하여 값 입력 후 문서 생성

## 설치 없이 바로 실행하기 (Windows)

1. 이 저장소 페이지에서 **Code → Download ZIP**으로 압축 파일을 내려받습니다.
2. 원하는 폴더에 압축을 풉니다.
3. 압축 푼 폴더 안의 **`run.bat`** 파일을 더블클릭합니다.
   - 최초 실행 시 필요한 패키지를 자동으로 설치합니다 (인터넷 연결 필요, 몇 분 정도 소요).
   - 이후 실행부터는 이미 설치된 패키지를 빠르게 확인만 하고 바로 프로그램이 열립니다 (보통 몇 초 이내). 업데이트로 새 패키지가 추가된 경우에도 자동으로 함께 설치됩니다.
   - 인터넷 연결이 없는 상태에서 실행하면, 이미 설치되어 있던 패키지 그대로 프로그램을 계속 엽니다.

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

- [xlsx (SheetJS)](https://github.com/SheetJS/sheetjs) — 엑셀/CSV 읽기(Extract) (Apache-2.0)
- [exceljs](https://github.com/exceljs/exceljs) — 서식이 적용된 엑셀 파일 쓰기(Load) (MIT)
- [docxtemplater](https://github.com/open-xml-templating/docxtemplater) — Word 템플릿 메일머지 (MIT)
- [pizzip](https://github.com/open-xml-templating/pizzip) — zip/OOXML 처리 (MIT)
- [iconv-lite](https://github.com/ashtuchkin/iconv-lite) — CSV 인코딩 변환 (MIT)
