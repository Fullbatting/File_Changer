@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
rem 사회사업실 도우미 실행: 두 번 누르면 PC에 없는 것(Node.js·패키지·실행 엔진)을 내려받아 준비하고(처음·업데이트 때만)
rem 프로그램을 띄운다. 프로그램 화면이 정상으로 뜨면 이 창은 저절로 닫힌다. 문제가 생기면 오류 창으로 알린다.
rem Node.js 찾기·내려받기는 src\launcher\node-env.bat, 그다음 순서는 src\launcher\launch.js에 있다.
rem 준비를 마친 폴더는 인터넷도 Node.js도 없는 PC로 통째로 옮겨도 그 안의 실행 엔진으로 돈다.

call "%~dp0src\launcher\node-env.bat"
if errorlevel 1 exit /b 1

"%SWH_NODE%" "%~dp0src\launcher\launch.js"
exit /b %errorlevel%
