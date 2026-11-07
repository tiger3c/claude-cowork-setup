@echo off
rem Install only Claude Code (terminal) for the current user. No admin needed.
setlocal
if not exist "%~dp0setup.ps1" (
  echo.
  echo  [ERROR] setup.ps1 not found. Extract the ZIP first ^(right-click - Extract All^).
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" -ClaudeCodeOnly
if errorlevel 1 if not errorlevel 2 pause
