@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
rem 사회사업실 도우미: 처음 설치·업데이트 때 쓰는 준비 파일.
rem 평소에는 "사회사업실 도우미.vbs"를 두 번 누르면 이 창 없이 바로 실행된다.
rem 준비가 끝나면 프로그램을 띄우고 이 창은 바로 닫는다. 문제가 생기면 오류 창으로 알린다.

echo ============================================
echo  사회사업실 도우미 - 실행 준비
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    set "FAIL_MSG=Node.js가 설치되어 있지 않습니다. https://nodejs.org/ 에서 Node.js LTS 버전을 설치한 뒤 다시 실행하세요."
    goto fail
)

rem 프로그램 실행 엔진(Electron 44)을 설치하려면 Node.js 22.12 이상이 필요하다
node -e "const [a, b] = process.versions.node.split('.').map(Number); process.exit(a > 22 || (a === 22 && b >= 12) ? 0 : 1)"
if errorlevel 1 (
    for /f %%v in ('node -v') do set "NODE_VER=%%v"
    goto old_node
)

if exist "node_modules" (
    echo [1/3] 필요한 패키지를 확인합니다. 이미 설치되어 있으면 몇 초 내로 끝납니다.
) else (
    echo [1/3] 필요한 패키지를 설치합니다. 인터넷 연결이 필요하며 최초 1회만 몇 분 정도 걸립니다.
)
echo.
call npm install
if errorlevel 1 (
    if not exist "node_modules" (
        set "FAIL_MSG=패키지 설치에 실패했습니다. 인터넷 연결 상태를 확인한 뒤 다시 실행하세요."
        goto fail
    )
    echo [경고] 패키지 확인 중 오류가 발생했습니다. 인터넷 연결 문제일 수 있어 기존 설치 상태로 계속 진행합니다.
)

rem Electron 42부터는 패키지 설치 때 실행 파일을 내려받지 않으므로 여기서 직접 받는다(이미 있으면 바로 끝난다).
echo.
echo [2/3] 프로그램 실행 엔진^(Electron^)을 확인합니다. 최초 1회는 내려받느라 몇 분 걸릴 수 있습니다.
node node_modules\electron\install.js
if errorlevel 1 (
    set "FAIL_MSG=프로그램 실행 엔진을 내려받지 못했습니다. 인터넷 연결 상태를 확인한 뒤 다시 실행하세요."
    goto fail
)
if not exist "node_modules\electron\dist\electron.exe" (
    set "FAIL_MSG=프로그램 실행 엔진(electron.exe)을 찾을 수 없습니다. run.bat을 다시 실행하세요."
    goto fail
)

rem 준비가 끝났다는 표시(바로 실행 파일이 다음부터 준비 단계를 건너뛰는 기준)
copy /y "package.json" "node_modules\.ready-package.json" > nul

echo.
echo [3/3] 프로그램을 실행합니다. 이 창은 곧 닫힙니다.
start "" "%~dp0node_modules\electron\dist\electron.exe" "%~dp0."
exit /b 0

:old_node
set "FAIL_MSG=설치된 Node.js 버전(%NODE_VER%)이 너무 낮습니다. 22.12 이상이 필요합니다. https://nodejs.org/ 에서 최신 LTS 버전을 설치한 뒤 다시 실행하세요."

:fail
echo.
echo [오류] %FAIL_MSG%
powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show($env:FAIL_MSG, '사회사업실 도우미 - 오류', 'OK', 'Error')"
exit /b 1
