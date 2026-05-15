@echo off
chcp 65001 > nul
title 엑셀 마스킹 도구 - 자동 빌드

echo.
echo ================================================
echo   엑셀 개인정보 마스킹 도구 - 자동 빌드 시작
echo ================================================
echo.

:: Python 설치 확인
python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo 아래 주소에서 Python을 설치해 주세요:
    echo https://www.python.org/downloads/
    echo.
    echo 설치 시 반드시 "Add Python to PATH" 체크!
    pause
    exit /b 1
)

echo [1/4] Python 확인 완료
python --version

echo.
echo [2/4] 필요한 패키지 설치 중... (잠시 기다려 주세요)
pip install pandas openpyxl pyinstaller --quiet --upgrade
if errorlevel 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
)
echo 패키지 설치 완료!

echo.
echo [3/4] EXE 파일 빌드 중... (1~2분 소요)
pyinstaller --onefile --windowed --name "엑셀_마스킹_도구" excel_masking.py --noconfirm
if errorlevel 1 (
    echo [오류] 빌드 실패. 오류 메시지를 확인해 주세요.
    pause
    exit /b 1
)

echo.
echo [4/4] 완료! EXE 파일 위치 확인 중...

if exist "dist\엑셀_마스킹_도구.exe" (
    echo.
    echo ================================================
    echo   빌드 성공!!
    echo   위치: %CD%\dist\엑셀_마스킹_도구.exe
    echo ================================================
    echo.
    echo 지금 바로 실행할까요? (dist 폴더가 열립니다)
    explorer dist
) else (
    echo [오류] EXE 파일을 찾을 수 없습니다.
)

echo.
pause
