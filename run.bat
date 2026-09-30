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

if exist "node_modules" (
    echo [1/2] 필요한 패키지를 확인합니다. 이미 설치되어 있으면 몇 초 내로 끝납니다.
) else (
    echo [1/2] 필요한 패키지를 설치합니다. 인터넷 연결이 필요하며 최초 1회만 몇 분 정도 걸립니다.
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
    echo [1/2] 패키지 확인/설치 완료.
)

echo.
echo [2/2] 프로그램을 실행합니다...
echo.
call npm start

echo.
echo 프로그램이 종료되었습니다.
pause
