@echo off
REM Downloads/caches the embedding model so that later runs work offline.
REM Run once after setup-venv.cmd.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Virtual environment not found. Run setup-venv.cmd first.
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
"%VENV_DIR%\Scripts\python.exe" -m app.main --check
endlocal
