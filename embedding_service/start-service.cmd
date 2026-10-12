@echo off
REM Starts the local semantic search service on 127.0.0.1.
REM The service is only reachable from this machine.
REM
REM Keep this window open while using semantic search in Outlook.
REM Close the window (or press Ctrl+C) to stop the service.
REM
REM Optional: set OAA_DEVICE=cuda to use an NVIDIA GPU if available.
REM           set OAA_MODEL_ID to override the model.
REM The service is incompatible with float16 on every backend.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Virtual environment not found.
  echo Run setup-venv.cmd first, then download-model.cmd, then this script.
  echo.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"

echo Starting the local semantic search service on 127.0.0.1:8765
echo Keep this window open. Press Ctrl+C to stop.
echo.

"%VENV_DIR%\Scripts\python.exe" -m app.main %*

echo.
echo The service has stopped.
pause
endlocal
