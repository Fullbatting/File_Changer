@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"

echo ============================================
echo  문서 변환 및 메일머지 자동화 프로그램
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo [오류] Node.js가 설치되어 있지 않습니다.
    echo 아래 주소에서 Node.js LTS 버전을 설치한 뒤 이 파일을 다시 실행하세요.
    echo https://nodejs.org/
    echo.
    pause
    exit /b 1
)

rem 프로그램 실행 엔진(Electron 44)을 설치하려면 Node.js 22.12 이상이 필요하다
node -e "const [a, b] = process.versions.node.split('.').map(Number); process.exit(a > 22 || (a === 22 && b >= 12) ? 0 : 1)"
if errorlevel 1 (
    for /f %%v in ('node -v') do echo [오류] 설치된 Node.js 버전^(%%v^)이 너무 낮습니다. 22.12 이상이 필요합니다.
    echo 아래 주소에서 최신 Node.js LTS 버전을 설치한 뒤 이 파일을 다시 실행하세요.
    echo https://nodejs.org/
    echo.
    pause
    exit /b 1
)

if exist "node_modules" (
    echo [1/3] 필요한 패키지를 확인합니다. 이미 설치되어 있으면 몇 초 내로 끝납니다.
) else (
    echo [1/3] 필요한 패키지를 설치합니다. 인터넷 연결이 필요하며 최초 1회만 몇 분 정도 걸립니다.
)
echo.
call npm install
if errorlevel 1 (
    if exist "node_modules" (
        echo.
        echo [경고] 패키지 확인 중 오류가 발생했습니다. 인터넷 연결 문제일 수 있어 기존 설치 상태로 계속 진행합니다.
        echo.
    ) else (
        echo.
        echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결 상태를 확인한 뒤 다시 시도하세요.
        echo.
        pause
        exit /b 1
    )
) else (
    echo.
    echo [1/3] 패키지 확인/설치 완료.
)

rem Electron 42부터는 패키지 설치 때 실행 파일을 내려받지 않으므로 여기서 직접 받는다.
rem 이미 같은 버전이 있으면 바로 끝난다.
echo.
echo [2/3] 프로그램 실행 엔진^(Electron^)을 확인합니다. 최초 1회는 내려받느라 몇 분 걸릴 수 있습니다.
node node_modules\electron\install.js
if errorlevel 1 (
    echo.
    echo [오류] 프로그램 실행 엔진을 내려받지 못했습니다. 인터넷 연결 상태를 확인한 뒤 다시 시도하세요.
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] 프로그램을 실행합니다...
echo.
call npm start

echo.
echo 프로그램이 종료되었습니다.
pause
