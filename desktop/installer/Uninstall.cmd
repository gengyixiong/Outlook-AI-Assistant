@echo off
setlocal
REM Require the user to extract the ZIP first. Running directly from inside a
REM ZIP archive extracts to a temporary folder that may be incomplete, which
REM makes PowerShell fail with "uninstall.ps1 does not exist".
if not exist "%~dp0uninstall.ps1" (
  echo.
  echo ERROR: uninstall.ps1 was not found next to this script.
  echo.
  echo You are probably running Uninstall.cmd from inside the ZIP file.
  echo Please EXTRACT the whole ZIP to a normal folder first
  echo ^(right-click the ZIP ^> Extract All^), then run Uninstall.cmd there.
  echo.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"
set UNINSTALL_RESULT=%ERRORLEVEL%
if not "%UNINSTALL_RESULT%"=="0" (
  echo.
  echo Uninstallation FAILED ^(exit code %UNINSTALL_RESULT%^).
  echo Review the message above before closing this window.
) else (
  echo.
  echo Uninstallation completed.
)
echo.
pause
exit /b %UNINSTALL_RESULT%
