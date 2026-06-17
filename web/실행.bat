@echo off
chcp 65001 >nul
title N2SF 마스킹 웹서버
cd /d "%~dp0"

set PY=python
where py >nul 2>nul && set PY=py

echo [1/2] 필요한 패키지 확인/설치 중...
%PY% -m pip install -r requirements.txt --quiet

echo [2/2] 서버 시작! 아래 표시되는 주소를 브라우저에 입력하세요.
echo        (종료하려면 이 창에서 Ctrl + C)
echo.
%PY% app.py
pause
