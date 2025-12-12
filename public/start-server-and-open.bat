@echo off
cd /d "%~dp0\.."
echo ===== 서버 실행 중... =====
start cmd /k "npm run server"

timeout /t 3 > nul

echo ===== 브라우저 열기 =====
start http://localhost:3000

echo.
echo 서버가 실행되었습니다!
echo 브라우저가 자동으로 열립니다.
echo.
pause




