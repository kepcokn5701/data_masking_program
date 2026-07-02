@echo off
chcp 65001 >nul
title 엑셀파일 개인정보 마스킹 - 빌드
cd /d "%~dp0"

set PY=python
where py >nul 2>nul && set PY=py

echo ============================================
echo   엑셀파일 개인정보 마스킹 - EXE/설치파일 빌드
echo ============================================
echo.
echo [1/4] 필요한 패키지 설치/확인...
%PY% -m pip install openpyxl xlrd pandas lxml html5lib beautifulsoup4 flask waitress pystray pillow pyinstaller --quiet --upgrade
if errorlevel 1 ( echo [오류] 패키지 설치 실패 - 인터넷/사내미러 확인 & pause & exit /b 1 )

echo.
echo [2/4] 프로그램(트레이+우클릭 통합) EXE 빌드... (1~3분)
%PY% -m PyInstaller --onefile --windowed --name "exel_info_masking_program(N2SF)" ^
  --paths "." --hidden-import pystray._win32 ^
  --hidden-import mask_cli --hidden-import register_context_menu --hidden-import uninstall ^
  --hidden-import make_manual --hidden-import manual_assets ^
  --hidden-import manual_assets_static --hidden-import manual_assets_lottie --hidden-import manual_assets_shots ^
  --hidden-import lxml --hidden-import html5lib --hidden-import bs4 ^
  --exclude-module matplotlib --exclude-module scipy ^
  --noconfirm main.py
if errorlevel 1 ( echo [오류] 프로그램 빌드 실패 & pause & exit /b 1 )

echo.
echo [3/4] 웹서버 EXE 빌드... (1~3분)
%PY% -m PyInstaller --onefile --console --name "exel_info_masking_webserver(N2SF)" ^
  --add-data "web/templates;templates" --paths "." --paths "web" ^
  --exclude-module matplotlib --exclude-module scipy --exclude-module PIL ^
  --noconfirm web/app.py
if errorlevel 1 ( echo [오류] 웹서버 빌드 실패 & pause & exit /b 1 )

echo.
echo [4/4] 설치파일(setup.exe) 생성... (Inno Setup 필요)
set ISCC=
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if defined ISCC (
  "%ISCC%" installer.iss
) else (
  echo [건너뜀] Inno Setup 미설치 - 설치파일은 생략됨(포터블 EXE는 dist에 있음^)
  echo          설치: winget install --id JRSoftware.InnoSetup --source winget
)

echo.
echo ============================================
echo   완료!
echo    - dist\exel_info_masking_program(N2SF).exe   (포터블 프로그램)
echo    - dist\exel_info_masking_webserver(N2SF).exe (웹서버)
echo    - installer_output\엑셀파일_개인정보_마스킹_설치.exe (설치파일)
echo ============================================
explorer dist
pause
