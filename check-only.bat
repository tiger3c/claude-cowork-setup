@echo off
rem Check this PC for Claude Cowork without changing anything (no admin needed).
setlocal
if not exist "%~dp0setup.ps1" (
  echo.
  echo  [ERROR] setup.ps1 not found. Extract the ZIP first ^(right-click - Extract All^).
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" -CheckOnly
echo.
pause
