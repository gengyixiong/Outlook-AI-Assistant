@echo off
REM Downloads and caches the embedding model so later runs work offline.
REM Run this ONCE after setup-venv.cmd and BEFORE building any index.
REM
REM On a machine that has not downloaded the model, indexing would otherwise
REM fail for every mail. This step requires network access (about 1.5 GB).
REM Afterwards the service runs fully offline.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Virtual environment not found. Run setup-venv.cmd first.
  exit /b 1
)

cd /d "%SCRIPT_DIR%"
echo Downloading the local semantic model. This can take several minutes...
"%VENV_DIR%\Scripts\python.exe" -m app.main --download-model
if errorlevel 1 (
  echo.
  echo Model download FAILED. Check the network or proxy and try again.
  echo If the repository requires a Hugging Face token, set HF_TOKEN first.
  exit /b 1
)

echo.
echo Model is ready. You can now start the service and build the index.
endlocal
