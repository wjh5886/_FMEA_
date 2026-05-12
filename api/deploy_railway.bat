@echo off
title Railway Deploy - FMEA API
cd /d E:\claude\FMEA\api

echo ================================
echo  FMEA API - Railway 배포
echo ================================
echo.

echo [1/3] Railway 로그인 (브라우저가 열립니다)...
railway login
echo 로그인 결과: %errorlevel%
pause

echo.
echo [2/3] Railway 프로젝트 초기화...
railway init
echo 초기화 결과: %errorlevel%
pause

echo.
echo [3/3] 배포 중...
railway up --detach
echo 배포 결과: %errorlevel%

echo.
railway open
pause
