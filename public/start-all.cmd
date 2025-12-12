@echo off
chcp 65001 > nul
title 개인정산 시스템 자동 실행
cd /d "%~dp0\.."

echo ========================================
echo 개인정산 시스템 서버 시작
echo ========================================
echo.

REM 기존 서버 프로세스 확인 및 종료
echo 기존 서버 프로세스 확인 중...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":3000" ^| findstr "LISTENING"') do (
    echo 기존 서버 프로세스 발견: %%a
    taskkill /F /PID %%a > nul 2>&1
    echo 기존 서버 프로세스를 종료했습니다.
    timeout /t 1 > nul
)

echo.
echo [1/3] 서버 실행 중...
echo 서버 창이 열립니다. 서버가 시작되는 동안 잠시만 기다려주세요...
start "개인정산 서버" cmd /k "title 개인정산 서버 && npm run server"

echo.
echo 서버가 시작되는 동안 대기 중... (최대 60초)
set /a count=0
set /a max_wait=30
:wait_server
timeout /t 2 > nul
set /a count+=1
echo 서버 상태 확인 중... (%count%/%max_wait%)

REM PowerShell을 사용하여 서버 상태 확인
powershell -Command "try { $response = Invoke-WebRequest -Uri 'http://localhost:3000/api/health' -TimeoutSec 2 -UseBasicParsing -ErrorAction Stop; if ($response.StatusCode -eq 200) { Write-Host 'OK'; exit 0 } else { exit 1 } } catch { exit 1 }" > nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo ✅ 서버가 정상적으로 시작되었습니다!
    goto server_ready
)

if %count% geq %max_wait% (
    echo.
    echo ⚠️ 서버 시작 시간이 오래 걸리고 있습니다. (60초 경과)
    echo 서버 창을 확인하여 오류가 있는지 확인해주세요.
    echo 계속 대기합니다... (서버가 시작되면 자동으로 진행됩니다)
    set /a count=0
)
goto wait_server

:server_ready
echo.
echo [2/3] 엑셀 변경 감시 시작...
start "엑셀 감시" cmd /k "title 엑셀 변경 감시 && npm run watch"

timeout /t 2 > nul

echo.
echo [3/3] 브라우저 열기...
timeout /t 1 > nul
start http://localhost:3000

echo.
echo ========================================
echo ✅ 모든 시스템이 정상적으로 시작되었습니다!
echo ========================================
echo.

REM 네트워크 IP 주소 가져오기 (PowerShell 사용)
echo 네트워크 IP 주소 확인 중...
for /f "tokens=*" %%i in ('powershell -Command "Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias ^(Get-NetAdapter ^| Where-Object Status -eq 'Up'^).Name | Where-Object IPAddress -notlike '127.*' | Select-Object -First 1 -ExpandProperty IPAddress"') do set "NETWORK_IP=%%i"

if defined NETWORK_IP (
    echo.
    echo 📍 접속 주소:
    echo    - 로컬: http://localhost:3000
    echo    - 네트워크: http://%NETWORK_IP%:3000
    echo.
    echo 💡 다른 사람과 공유하려면 네트워크 주소를 사용하세요!
    echo    같은 네트워크에 연결된 다른 기기에서 접속 가능합니다.
    echo    예: http://%NETWORK_IP%:3000
    echo.
) else (
    echo.
    echo 📍 접속 주소: http://localhost:3000
    echo.
    echo 💡 네트워크 IP를 확인할 수 없습니다.
    echo    서버 창에서 네트워크 주소를 확인하세요.
    echo.
)
echo 백엔드 서버: http://localhost:3000
echo 프론트 화면이 자동으로 열립니다.
echo.
echo ⚠️ 중요: 브라우저 주소창이 http://localhost:3000 인지 확인하세요!
echo    다른 주소(예: public/index.html, 포트 5500 등)로 열리면 오류가 발생합니다.
echo.
echo 서버를 종료하려면 "개인정산 서버" 창에서 Ctrl+C를 누르세요.
echo.
pause
