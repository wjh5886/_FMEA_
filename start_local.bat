@echo off
chcp 65001 > nul
title FMEA 로컬 서버

echo ========================================
echo  FMEA 로컬 서버 시작
echo ========================================
echo.

:: Python 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    pause & exit /b 1
)

:: Node.js 확인
node --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Node.js가 설치되지 않았습니다.
    pause & exit /b 1
)

:: 의존성 설치 (첫 실행 시)
if not exist "%~dp0api\__pycache__" (
    echo [설치] Python 패키지 설치 중...
    pip install -r "%~dp0api\requirements.txt" -q
)
if not exist "%~dp0fmea-web\node_modules" (
    echo [설치] Node 패키지 설치 중...
    cd /d "%~dp0fmea-web" && npm install
)

:: 내 IP 주소 출력
echo.
echo ── 내 PC의 IP 주소 ─────────────────────
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set IP=%%a
    setlocal enabledelayedexpansion
    set IP=!IP: =!
    echo   !IP!
    endlocal
)
echo ─────────────────────────────────────────
echo.
echo  팀원 접속 주소: http://[위 IP]:3000
echo  예) http://192.168.1.100:3000
echo.

:: FastAPI 백엔드 시작 (새 창)
start "FMEA Backend (포트 8000)" cmd /k "cd /d %~dp0api && uvicorn main:app --host 0.0.0.0 --port 8000 --reload"

:: 잠시 대기 후 Next.js 시작 (새 창)
timeout /t 2 /nobreak > nul
start "FMEA Frontend (포트 3000)" cmd /k "cd /d %~dp0fmea-web && npm run dev -- --hostname 0.0.0.0"

echo.
echo  두 서버가 시작되었습니다.
echo  브라우저에서 http://localhost:3000 으로 접속하세요.
echo  팀원은 위 IP 주소로 접속 가능합니다.
echo.
echo  종료하려면 각 창을 닫으세요.
echo ========================================
pause
