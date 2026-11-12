@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title 웹 메뉴 수집기 - 웹 화면
if not exist ".venv\Scripts\python.exe" call "%~dp0install.bat"
echo 기본 브라우저에 실행 화면이 열립니다. 이 창을 닫으면 프로그램이 종료됩니다.
".venv\Scripts\python.exe" launcher.py %*
if errorlevel 1 pause
