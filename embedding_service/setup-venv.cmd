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
REM IMPORTANT: Python 3.10-3.12 is required. PyTorch (and several other
REM dependencies) do not publish prebuilt wheels for Python 3.13+, so on a
REM newer interpreter pip tries to compile from source and fails with
REM "Microsoft Visual C++ 14.0 or greater is required".
REM
REM Python discovery order:
REM   1. the "py -3.12" / "-3.11" / "-3.10" launcher (python.org installs)
REM   2. a supported python.exe on PATH, skipping the Microsoft Store stub
REM   3. common per-user / per-machine install locations for 3.12/3.11/3.10

setlocal EnableDelayedExpansion
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%.venv
set PYTHON_EXE=
set PYTHON_ARGS=
set NO_PAUSE=
if /i "%~1"=="--no-pause" set NO_PAUSE=1

REM --- 1. the py launcher, preferring a supported version ----------------
where py >nul 2>nul
if not errorlevel 1 (
  for %%v in (3.12 3.11 3.10) do (
    if not defined PYTHON_EXE (
      py -%%v -c "import sys" >nul 2>nul
      if not errorlevel 1 (
        set "PYTHON_ARGS=-%%v"
        for /f "delims=" %%p in ('py -%%v -c "import sys;print(sys.executable)"') do set "PYTHON_EXE=%%p"
      )
    )
  )
)

REM --- 2. a supported python.exe on PATH (skip the Store stub) -----------
if not defined PYTHON_EXE (
  for /f "delims=" %%p in ('where python 2^>nul') do (
    if not defined PYTHON_EXE (
      echo %%p | findstr /i "WindowsApps" >nul
      if errorlevel 1 (
        "%%p" -c "import sys;sys.exit(0 if sys.version_info[:2] in ((3,10),(3,11),(3,12)) else 1)" >nul 2>nul
        if not errorlevel 1 set "PYTHON_EXE=%%p"
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
  echo A supported Python ^(3.10, 3.11 or 3.12^) was not found.
  echo ============================================================
  echo.
  echo Installed Python versions reported by the py launcher:
  echo.
  py -0 2>nul
  if errorlevel 1 echo   ^(the py launcher is not available^)
  echo.
  echo Python 3.13 and newer cannot be used: PyTorch does not publish
  echo prebuilt wheels for them yet, so installation fails while
  echo compiling from source.
  echo.
  echo Install Python 3.12 from:
  echo   https://www.python.org/downloads/release/python-3129/
  echo Tick 'Add python.exe to PATH' during setup, then re-run this script.
  echo.
  echo Do NOT install the Microsoft Store version of Python; it is a
  echo stub. If 'python' opens the Store, turn off the python.exe and
  echo python3.exe aliases under:
  echo   Settings ^> Apps ^> Advanced app settings ^> App execution aliases
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

REM Final safety check: refuse an unsupported interpreter explicitly.
"%PYTHON_EXE%" -c "import sys;sys.exit(0 if sys.version_info[:2] in ((3,10),(3,11),(3,12)) else 1)"
if errorlevel 1 (
  echo.
  echo This Python is not supported. Use Python 3.10-3.12.
  echo PyTorch has no prebuilt wheel for 3.13+ and the install would try
  echo to compile from source and fail.
  echo.
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
echo If the error mentions "Microsoft Visual C++ 14.0 or greater is
echo required", the interpreter is too new. Install Python 3.12 and retry.
echo.
if not defined NO_PAUSE pause
endlocal
exit /b 1
