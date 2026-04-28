@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%Install-NewUserAutomation.ps1"

if not exist "%SCRIPT_PATH%" (
  echo Installer script not found:
  echo %SCRIPT_PATH%
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
set "EXITCODE=%ERRORLEVEL%"

if not "%EXITCODE%"=="0" (
  echo.
  echo Installer exited with code %EXITCODE%.
)

pause
exit /b %EXITCODE%
