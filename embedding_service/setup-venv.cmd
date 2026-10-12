@echo off
REM Sets up the Python virtual environment for the semantic search service.
REM Usage: setup-venv.cmd
REM
REM Creates .venv next to this script and installs the CPU build of PyTorch
REM (no CUDA required), then the remaining dependencies. Run once.
REM
REM The window stays open at the end so you can read the result.
REM Pass --no-pause to suppress the final pause (used by install-service.ps1).
REM
REM Python discovery order:
REM   1. the "py" launcher (installed by python.org installers)
REM   2. a real python.exe on PATH, skipping the Microsoft Store alias stub
REM   3. common per-user / per-machine install locations
REM The Store alias is deliberately rejected because it is a stub that only
REM opens the Store page and cannot create a virtual environment.

setlocal EnableDelayedExpansion
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv
set PYTHON_EXE=
set NO_PAUSE=
if /i "%~1"=="--no-pause" set NO_PAUSE=1

REM --- 1. the py launcher -------------------------------------------------
where py >nul 2>nul
if not errorlevel 1 (
  for /f "delims=" %%v in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
    set "PYTHON_EXE=%%v"
  )
)

REM --- 2. a real python.exe on PATH (skip the Store stub) -----------------
if not defined PYTHON_EXE (
  for /f "delims=" %%p in ('where python 2^>nul') do (
    if not defined PYTHON_EXE (
      echo %%p | findstr /i "WindowsApps" >nul
      if errorlevel 1 (
        set "PYTHON_EXE=%%p"
      )
    )
  )
)

REM --- 3. common install locations ---------------------------------------
if not defined PYTHON_EXE (
  for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%LOCALAPPDATA%\Programs\Python\Python311"
    "%LOCALAPPDATA%\Programs\Python\Python310"
    "%ProgramFiles%\Python312"
    "%ProgramFiles%\Python311"
    "%ProgramFiles%\Python310"
  ) do (
    if not defined PYTHON_EXE (
      if exist "%%~d\python.exe" set "PYTHON_EXE=%%~d\python.exe"
    )
  )
)

if not defined PYTHON_EXE (
  echo ============================================================
  echo Python 3.10-3.12 was not found.
  echo ============================================================
  echo.
  echo A common cause is that "python" points at the Microsoft Store
  echo stub instead of a real install. To fix it:
  echo.
  echo   1. Install Python from https://www.python.org/downloads/
  echo      and tick "Add python.exe to PATH" during setup.
  echo   2. If it is already installed but still not found, open
  echo      Settings ^> Apps ^> Advanced app settings ^>
  echo      App execution aliases, and turn OFF the "python.exe" and
  echo      "python3.exe" aliases.
  echo   3. Re-run this script.
  echo.
  if not defined NO_PAUSE pause
  exit /b 1
)

echo Using Python: %PYTHON_EXE%
"%PYTHON_EXE%" -c "import sys;print(sys.version)"
if errorlevel 1 (
  echo That Python could not be started.
  if not defined NO_PAUSE pause
  exit /b 1
)

if not exist "%VENV_DIR%" (
  echo Creating virtual environment at %VENV_DIR%
  "%PYTHON_EXE%" -m venv "%VENV_DIR%"
  if errorlevel 1 (
    echo.
    echo Failed to create the virtual environment.
    echo If the message above mentions the Microsoft Store, disable the
    echo python.exe / python3.exe aliases under Settings ^> Apps ^>
    echo Advanced app settings ^> App execution aliases, then retry.
    echo.
    if not defined NO_PAUSE pause
    exit /b 1
  )
)

call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
  echo Failed to activate the virtual environment.
  if not defined NO_PAUSE pause
  exit /b 1
)

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
if not defined NO_PAUSE pause
endlocal
exit /b 0

:failed
echo.
echo Setup FAILED. Review the messages above.
echo.
if not defined NO_PAUSE pause
endlocal
exit /b 1
