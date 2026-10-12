@echo off
REM Collects diagnostics for the local semantic search service and the
REM add-in, then writes a report to the Desktop.
REM
REM Double-click this file. When it finishes, the report path is printed.
REM Send the contents of that report back for analysis.

setlocal
set SCRIPT_DIR=%~dp0

echo ============================================================
echo Outlook AI Assistant - diagnostic report
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass ^
  -File "%SCRIPT_DIR%diagnose.ps1" %*
set RESULT=%ERRORLEVEL%

echo.
if not "%RESULT%"=="0" (
  echo Diagnostic script exited with code %RESULT%.
)
echo.
pause
endlocal
exit /b %RESULT%
