@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 웹 메뉴 수집기
if not exist ".venv\Scripts\python.exe" (
  echo 처음 실행이므로 설치를 먼저 진행합니다.
  call "%~dp0설치.bat"
  if not exist ".venv\Scripts\python.exe" exit /b 1
)
".venv\Scripts\python.exe" -m menu_crawler %*
echo.
pause
