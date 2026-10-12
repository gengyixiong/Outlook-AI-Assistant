@echo off
REM Starts the local semantic search service on 127.0.0.1.
REM The service is only reachable from this machine.
REM
REM Optional: set OAA_DEVICE=cuda to use an NVIDIA GPU if available.
REM           set OAA_MODEL_ID to override the model.
REM The service is incompatible with float16 on every backend.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Virtual environment not found. Run setup-venv.cmd first.
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
"%VENV_DIR%\Scripts\python.exe" -m app.main %*
endlocal
