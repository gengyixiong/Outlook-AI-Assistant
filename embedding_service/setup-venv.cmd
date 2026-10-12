@echo off
REM Sets up the Python virtual environment for the semantic search service.
REM Usage: setup-venv.cmd
REM
REM Creates .venv next to this script and installs the CPU build of PyTorch
REM (no CUDA required), then the remaining dependencies. Run once.
REM
REM The window stays open at the end so you can read the result.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

where python >nul 2>nul
if errorlevel 1 (
  echo Python 3.10-3.12 was not found on PATH.
  echo Install it from https://www.python.org/downloads/ and retry.
  echo.
  pause
  exit /b 1
)

if not exist "%VENV_DIR%" (
  echo Creating virtual environment at %VENV_DIR%
  python -m venv "%VENV_DIR%"
  if errorlevel 1 (
    echo Failed to create the virtual environment.
    pause
    exit /b 1
  )
)

call "%VENV_DIR%\Scripts\activate.bat"

echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 goto :failed

echo.
echo Installing PyTorch (CPU build). This is the largest step...
python -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu
if errorlevel 1 goto :failed

echo.
echo Installing remaining dependencies...
python -m pip install -r "%SCRIPT_DIR%requirements.txt"
if errorlevel 1 goto :failed

echo.
echo Setup complete. Next step: run download-model.cmd
echo.
pause
endlocal
exit /b 0

:failed
echo.
echo Setup FAILED. Review the messages above.
echo.
pause
endlocal
exit /b 1
