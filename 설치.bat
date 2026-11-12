@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 웹 메뉴 수집기 - 설치

set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto nopython
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if errorlevel 1 goto nopython

echo [1/3] 실행 환경 만들기 .venv
if not exist ".venv\Scripts\python.exe" %PY% -m venv .venv
if errorlevel 1 goto fail

echo [2/3] 필요한 프로그램 설치
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q --upgrade pip
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto fail

echo [3/3] 브라우저 확인
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" goto browser_ok
if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" goto browser_ok
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" goto browser_ok
echo Edge/Chrome을 찾지 못해 내장 Chromium을 설치합니다.
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto fail
:browser_ok

echo.
echo 설치가 끝났습니다. 이제 실행.bat을 더블클릭하세요. (웹 화면이 편하면 화면실행.bat)
pause
exit /b 0

:nopython
echo Python 3.10 이상이 필요합니다.
echo https://www.python.org/downloads/ 에서 설치하고, 설치 화면의 "Add python.exe to PATH"를 체크하세요.
pause
exit /b 1

:fail
echo.
echo 설치 중 오류가 발생했습니다. 위 메시지를 확인하세요. 회사 네트워크라면 프록시·보안 프로그램 설정을 확인하세요.
pause
exit /b 1
