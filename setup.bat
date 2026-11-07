@echo off
rem Claude Cowork setup - double-click this file.
setlocal
if not exist "%~dp0setup.ps1" (
  echo.
  echo  [ERROR] setup.ps1 not found next to this file.
  echo  If you opened the ZIP directly, right-click the ZIP - "Extract All" first,
  echo  then run setup.bat from the extracted folder.
  echo.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" %*
rem exit code 1 = PowerShell could not run the script (policy etc.); show the message
if errorlevel 1 if not errorlevel 2 pause
