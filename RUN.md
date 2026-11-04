# 실행 방법

Node.js 22.12 이상이 필요합니다. Windows에서는 저장소 폴더의 `run.bat`을 실행합니다.

macOS 또는 Linux에서는 저장소 폴더에서 다음 명령을 실행합니다.

```bash
npm install
node node_modules/electron/install.js
npm start
```

테스트는 `npm test`로 실행합니다. 디스플레이가 없는 Linux 서버에서는 `xvfb-run -a npm test`를 사용합니다. 자세한 내용은 [README](README.md)를 참고하세요.
