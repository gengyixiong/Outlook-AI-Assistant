@echo off
REM Installs the local semantic search service to a stable per-user location.
REM
REM Run this ONCE after extracting the release. It copies the service to
REM   %LOCALAPPDATA%\OutlookAiAssistant\embedding_service
REM so that the extracted folder can be deleted afterwards.
REM
REM The window stays open at the end so you can read the result.

setlocal
echo ============================================================
echo Installing the local semantic search service
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-service.ps1" %*
set RESULT=%ERRORLEVEL%

echo.
if not "%RESULT%"=="0" (
  echo Installation FAILED. Review the messages above.
) else (
  echo Installation completed.
)
echo.
pause
endlocal
exit /b %RESULT%
