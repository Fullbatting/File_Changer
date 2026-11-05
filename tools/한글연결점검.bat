@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0.."

echo ============================================
echo  사회사업실 도우미 - 한글 연결 점검
echo ============================================
echo  한글(한컴오피스)이 이 PC에서 자동으로 서식을 열고 저장할 수 있는지 시험합니다.
echo  시험할 서식(.hwp/.hwpx)이 있으면 이 파일 위로 끌어다 놓아 실행하세요.
echo  시험 중 한글에서 "접근 허용" 창이 뜨면 [모두 허용] 또는 [허용]을 누르세요.
echo.

call "%~dp0..\src\launcher\node-env.bat"
if errorlevel 1 exit /b 1

if not exist "node_modules" (
    echo 필요한 패키지를 설치합니다. 인터넷 연결이 필요하며 최초 1회만 몇 분 걸립니다.
    call npm install
)

"%SWH_NODE%" tools\hancom-check.js %*

echo.
echo 점검이 끝났습니다. 위 결과 폴더의 파일을 보내 주세요.
pause
