@echo off
rem src\launcher\node-env.bat — run.bat 등이 call로 부른다(setlocal 없이, 부른 쪽의 환경을 바꾼다).
rem 이 프로그램의 시작 스크립트를 돌릴 Node.js를 찾아 SWH_NODE에 넣는다. 찾는 순서:
rem   1) 프로그램 폴더 안에 내려받아 둔 것(runtime\node)
rem   2) PC에 설치된 Node.js(22.12 이상일 때만)
rem   3) 이미 준비된 폴더면 그 안의 실행 엔진(Electron). Electron에는 Node.js가 들어 있어
rem      ELECTRON_RUN_AS_NODE=1로 Node.js처럼 쓸 수 있다. 준비된 폴더를 USB로 옮긴 인터넷 없는 PC는 여기서 끝난다
rem   4) 아무것도 없으면 공식 Node.js를 runtime\node에 내려받는다(get-node.ps1, 처음 한 번)
rem 성공하면 종료 코드 0. 실패하면 오류 창을 띄운 뒤(또는 이미 뜬 뒤) 종료 코드 1.

for %%I in ("%~dp0..\..") do set "SWH_ROOT=%%~fI"
set "SWH_NODE_DIR=%SWH_ROOT%\runtime\node"
set "SWH_ELECTRON=%SWH_ROOT%\node_modules\electron\dist\electron.exe"
set "ELECTRON_RUN_AS_NODE="

if exist "%SWH_NODE_DIR%\node.exe" goto use_local
call :system_node_ok
if errorlevel 1 goto try_electron
set "SWH_NODE=node"
exit /b 0

:try_electron
if not exist "%SWH_ELECTRON%" goto download
set "SWH_NODE=%SWH_ELECTRON%"
set "ELECTRON_RUN_AS_NODE=1"
exit /b 0

:download
echo 이 PC에 쓸 수 있는 Node.js(22.12 이상)가 없어 프로그램 폴더 안에 내려받습니다. 처음 한 번만 합니다.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0get-node.ps1" -Dest "%SWH_NODE_DIR%"
set "SWH_PS_CODE=%errorlevel%"
rem 3: 스크립트가 이미 오류 창을 띄웠다. 그 밖의 실패: 스크립트가 돌지 못했다
if "%SWH_PS_CODE%"=="3" exit /b 1
if not "%SWH_PS_CODE%"=="0" goto ps_failed
if not exist "%SWH_NODE_DIR%\node.exe" goto ps_failed

:use_local
set "PATH=%SWH_NODE_DIR%;%PATH%"
set "SWH_NODE=%SWH_NODE_DIR%\node.exe"
exit /b 0

:ps_failed
rem PowerShell이 스크립트를 아예 돌리지 못한 경우(회사 보안 정책 등). 스크립트의 오류 창이 뜨지 않았으므로 여기서 띄운다
powershell -NoProfile -Command "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show('Node.js를 자동으로 내려받지 못했습니다.' + [Environment]::NewLine + 'https://nodejs.org/ 에서 Node.js LTS 버전을 설치한 뒤 run.bat을 다시 실행하세요.', '사회사업실 도우미 - 오류', 'OK', 'Error')"
exit /b 1

:system_node_ok
where node >nul 2>nul
if errorlevel 1 exit /b 1
node -e "var v=process.versions.node.split('.').map(Number);process.exit(v[0]>22||(v[0]===22&&v[1]>=12)?0:1)" >nul 2>nul
exit /b %errorlevel%
