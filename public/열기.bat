@echo off
chcp 65001 > nul
cd /d "%~dp0\.."

echo ===== 서버 상태 확인 중... =====
curl -s http://localhost:3000/api/health > nul 2>&1
if %errorlevel% equ 0 (
    echo 서버가 이미 실행 중입니다.
    echo 브라우저를 엽니다...
    start http://localhost:3000
) else (
    echo 서버가 실행 중이지 않습니다. 서버를 시작합니다...
    echo.
    start cmd /k "npm run server"
    
    echo 서버가 시작되는 동안 잠시 기다립니다...
    timeout /t 5 > nul
    
    echo 브라우저를 엽니다...
    start http://localhost:3000
    
    echo.
    echo 서버가 실행되었습니다!
    echo 브라우저가 자동으로 열립니다.
    echo.
)

pause




