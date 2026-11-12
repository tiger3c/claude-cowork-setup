@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo 설치.bat을 먼저 실행하세요.
  pause
  exit /b 1
)
echo 내장 Chromium 브라우저를 설치합니다. Edge/Chrome을 쓸 수 없는 PC에서만 필요합니다.
".venv\Scripts\python.exe" -m playwright install chromium
pause
