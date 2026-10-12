@echo off
REM Downloads and caches the embedding model so later runs work offline.
REM Run this ONCE after setup-venv.cmd and BEFORE building any index.
REM
REM On a machine that has not downloaded the model, indexing would otherwise
REM fail for every mail. This step requires network access (about 1.5 GB).
REM Afterwards the service runs fully offline.
REM
REM The window stays open at the end so you can read the result.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo Virtual environment not found. Run setup-venv.cmd first.
  echo.
  pause
  exit /b 1
)

cd /d "%SCRIPT_DIR%"

echo ============================================================
echo Downloading the local semantic model (about 1.5 GB).
echo This can take several minutes. Progress is shown below.
echo Keep this window open until it reports SUCCESS.
echo ============================================================
echo.

"%VENV_DIR%\Scripts\python.exe" -m app.main --download-model
set RESULT=%ERRORLEVEL%

echo.
if not "%RESULT%"=="0" (
  echo Model download FAILED.
  echo Check the network or proxy and try again.
  echo If the repository requires a Hugging Face token, run:  set HF_TOKEN=...
) else (
  echo Model is ready. You can now run start-service.cmd
)
echo.
pause
endlocal
exit /b %RESULT%
