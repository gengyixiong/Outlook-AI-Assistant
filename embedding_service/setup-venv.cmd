@echo off
REM Sets up the Python virtual environment for the semantic search service.
REM Usage: setup-venv.cmd
REM
REM Creates .venv next to this script and installs the CPU build of PyTorch
REM (no CUDA required), then the remaining dependencies. Run once.

setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv

where python >nul 2>nul
if errorlevel 1 (
  echo Python 3.10-3.12 was not found on PATH.
  echo Install it from https://www.python.org/downloads/ and retry.
  exit /b 1
)

if not exist "%VENV_DIR%" (
  echo Creating virtual environment at %VENV_DIR%
  python -m venv "%VENV_DIR%"
  if errorlevel 1 exit /b 1
)

call "%VENV_DIR%\Scripts\activate.bat"

python -m pip install --upgrade pip
if errorlevel 1 exit /b 1

echo Installing PyTorch (CPU build)...
python -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cpu
if errorlevel 1 exit /b 1

echo Installing remaining dependencies...
python -m pip install -r "%SCRIPT_DIR%requirements.txt"
if errorlevel 1 exit /b 1

echo.
echo Setup complete. Start the service with start-service.cmd
endlocal
