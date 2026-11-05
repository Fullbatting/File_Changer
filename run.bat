@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
rem 사회사업실 도우미 실행: 두 번 누르면 준비(처음·업데이트 때만) 뒤 프로그램을 띄우고,
rem 프로그램 화면이 정상으로 뜨면 이 창은 저절로 닫힌다. 문제가 생기면 오류 창으로 알린다.
rem 실제 순서는 src\launcher\launch.js에 있다. 여기서는 그것을 돌릴 Node.js가 있는지만 본다.

where node >nul 2>nul
if errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [void][System.Windows.MessageBox]::Show('Node.js가 설치되어 있지 않습니다.' + [Environment]::NewLine + 'https://nodejs.org/ 에서 Node.js LTS 버전을 설치한 뒤 run.bat을 다시 실행하세요.', '사회사업실 도우미 - 오류', 'OK', 'Error')"
    exit /b 1
)

node "%~dp0src\launcher\launch.js"
exit /b %errorlevel%
