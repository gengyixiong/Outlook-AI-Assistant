@echo off
REM ===================================================================
REM  Outlook AI Assistant - one-click install / update
REM
REM  Double-click this file. It checks what is already installed and only
REM  performs what is missing or outdated:
REM    1. Python 3.10-3.12 check
REM    2. Outlook add-in (COM registration)
REM    3. Semantic search service code
REM    4. Python dependencies (venv)  -- only once
REM    5. Embedding model            -- only once
REM
REM  Upgrade keeps the venv, model cache, semantic index, settings and the
REM  encrypted API key. Safe to run any number of times.
REM
REM  Extract the ZIP first. Running from inside the archive fails because
REM  the temporary folder is incomplete.
REM ===================================================================

setlocal
set SCRIPT_DIR=%~dp0

if not exist "%SCRIPT_DIR%setup.ps1" (
  echo.
  echo ERROR: setup.ps1 was not found next to this file.
  echo.
  echo You are probably running from inside the ZIP file.
  echo Right-click the ZIP, choose "Extract All", then run setup.cmd
  echo from the extracted folder.
  echo.
  pause
  exit /b 1
)

if not exist "%SCRIPT_DIR%VERSION" (
  echo.
  echo ERROR: VERSION file is missing.
  echo Re-extract the complete release ZIP and try again.
  echo.
  pause
  exit /b 1
)

echo ============================================================
echo Outlook AI Assistant - install / update
echo ============================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass ^
  -File "%SCRIPT_DIR%setup.ps1" %*
set RESULT=%ERRORLEVEL%

if not "%RESULT%"=="0" (
  echo.
  echo Some steps did not complete. Review the messages above.
  echo You can run diagnose.cmd for a full report.
)

endlocal
exit /b %RESULT%
