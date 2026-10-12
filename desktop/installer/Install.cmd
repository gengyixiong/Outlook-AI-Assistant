@echo off
setlocal
REM Require the user to extract the ZIP first. Running directly from inside a
REM ZIP archive extracts to a temporary folder that may be incomplete, which
REM makes PowerShell fail with "install.ps1 does not exist".
if not exist "%~dp0install.ps1" (
  echo.
  echo ERROR: install.ps1 was not found next to this script.
  echo.
  echo You are probably running Install.cmd from inside the ZIP file.
  echo Please EXTRACT the whole ZIP to a normal folder first
  echo ^(right-click the ZIP ^> Extract All^), then run Install.cmd there.
  echo.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set INSTALL_RESULT=%ERRORLEVEL%
if not "%INSTALL_RESULT%"=="0" (
  echo.
  echo Installation FAILED ^(exit code %INSTALL_RESULT%^).
  echo Review the message above before closing this window.
) else (
  echo.
  echo Installation completed.
)
echo.
pause
exit /b %INSTALL_RESULT%

