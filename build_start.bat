@echo off
chcp 65001 >nul
title N2SF 마스킹 - EXE 빌드 (데스크톱 + 웹서버)
cd /d "%~dp0"

set PY=python
where py >nul 2>nul && set PY=py

echo ============================================
echo   N2SF 마스킹 EXE 빌드
echo ============================================
echo.
echo [1/3] 필요한 패키지 설치/확인...
%PY% -m pip install openpyxl xlrd pandas lxml flask waitress pyinstaller --quiet --upgrade
if errorlevel 1 ( echo [오류] 패키지 설치 실패 - 인터넷/사내미러 확인 & pause & exit /b 1 )

echo.
echo [2/3] 데스크톱 EXE 빌드... (1~3분)
%PY% -m PyInstaller --onefile --windowed --name "exel_info_masking_program(N2SF)" ^
  --exclude-module matplotlib --exclude-module scipy --exclude-module PIL ^
  --noconfirm excel_masking.py
if errorlevel 1 ( echo [오류] 데스크톱 빌드 실패 & pause & exit /b 1 )

echo.
echo [3/3] 웹서버 EXE 빌드... (1~3분)
%PY% -m PyInstaller --onefile --console --name "exel_info_masking_webserver(N2SF)" ^
  --add-data "web/templates;templates" --paths "." --paths "web" ^
  --exclude-module matplotlib --exclude-module scipy --exclude-module PIL ^
  --noconfirm web/app.py
if errorlevel 1 ( echo [오류] 웹서버 빌드 실패 & pause & exit /b 1 )

echo.
echo ============================================
echo   완료! dist 폴더에 EXE 2개 생성됨
echo    - exel_info_masking_program(N2SF).exe   (데스크톱)
echo    - exel_info_masking_webserver(N2SF).exe (웹서버)
echo ============================================
explorer dist
pause
