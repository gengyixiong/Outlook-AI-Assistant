# Installs the local semantic search service to a stable per-user location.
#
# WHY: the service code and its .venv used to live inside whatever folder the
# release ZIP was extracted to (often Downloads). Deleting that folder broke
# semantic search. This script copies the service to
#   %LOCALAPPDATA%\OutlookAiAssistant\embedding_service
# and creates the .venv there, so the extracted ZIP can be deleted afterwards.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\install-service.ps1
#
# Options:
#   -SkipVenv     copy the code only; run setup-venv later
#   -Force        reinstall even if already present

[CmdletBinding()]
param(
  [switch]$SkipVenv,
  [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$sourceRoot = $PSScriptRoot
$productRoot = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$targetRoot = Join-Path $productRoot "embedding_service"

Write-Host "Outlook AI Assistant - local semantic service installer"
Write-Host ""

if (-not (Test-Path -LiteralPath (Join-Path $sourceRoot "app\main.py"))) {
  throw "This script must be run from the embedding_service folder (app\main.py not found)."
}

# Refuse to install over ourselves (e.g. running it from the install dir).
$resolvedSource = (Resolve-Path -LiteralPath $sourceRoot).Path
$resolvedTarget = [System.IO.Path]::GetFullPath($targetRoot)
if ($resolvedSource.Equals($resolvedTarget, [System.StringComparison]::OrdinalIgnoreCase)) {
  Write-Host "Already running from the install location: $resolvedTarget"
  Write-Host "Nothing to copy."
}
else {
  if ((Test-Path -LiteralPath $targetRoot) -and -not $Force) {
    Write-Host "Existing installation found at:"
    Write-Host "  $targetRoot"
    Write-Host "Updating the code (your .venv and downloaded model are kept)."
  }

  New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null

  # Copy the service code. Deliberately never touch .venv or any cached model.
  $files = @(
    "README.md",
    "requirements.txt",
    "setup-venv.cmd",
    "download-model.cmd",
    "start-service.cmd"
  )
  foreach ($name in $files) {
    $source = Join-Path $sourceRoot $name
    if (Test-Path -LiteralPath $source) {
      Copy-Item -LiteralPath $source -Destination $targetRoot -Force
    }
    else {
      Write-Warning "Expected file is missing from the release: $name"
    }
  }
  # The uninstaller is not copied into the install dir; it runs from the
  # release folder against the install dir.

  $appTarget = Join-Path $targetRoot "app"
  if (Test-Path -LiteralPath $appTarget) {
    Remove-Item -LiteralPath $appTarget -Recurse -Force
  }
  Copy-Item -LiteralPath (Join-Path $sourceRoot "app") -Destination $appTarget -Recurse

  # Remove any bytecode caches so a stale .pyc can never shadow new code.
  Get-ChildItem -LiteralPath $targetRoot -Recurse -Directory `
    -Filter "__pycache__" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }

  Write-Host "Service code installed to:"
  Write-Host "  $targetRoot"
}

# Create (or reuse) the virtual environment in the install location.
$venvPath = Join-Path $targetRoot ".venv"
$venvPython = Join-Path $venvPath "Scripts\python.exe"

if ($SkipVenv) {
  Write-Host ""
  Write-Host "Skipped virtual environment creation (-SkipVenv)."
}
elseif (Test-Path -LiteralPath $venvPython) {
  Write-Host ""
  Write-Host "Virtual environment already present; keeping it."
}
else {
  Write-Host ""
  Write-Host "Creating the virtual environment. This installs PyTorch and can"
  Write-Host "take several minutes..."
  $setup = Join-Path $targetRoot "setup-venv.cmd"
  & cmd.exe /c "`"$setup`" --no-pause"
  if ($LASTEXITCODE -ne 0) {
    throw @"
setup-venv.cmd failed with exit code $LASTEXITCODE.

Most common causes:
  1. Python 3.13 or newer is installed. PyTorch has no prebuilt wheels for
     it, so the install tries to compile from source and fails with
     "Microsoft Visual C++ 14.0 or greater is required".
     Fix: install Python 3.12 from https://www.python.org/downloads/ and
     re-run install-service.cmd.
  2. No real Python on PATH (only the Microsoft Store stub). Turn off the
     python.exe / python3.exe aliases under Settings > Apps > Advanced app
     settings > App execution aliases.
"@
  }
}

Write-Host ""
Write-Host "Done. The extracted release folder can now be deleted."

# Create a Start Menu shortcut so the user never has to remember the path.
try {
  $startMenu = Join-Path $env:APPDATA `
    "Microsoft\Windows\Start Menu\Programs\Outlook AI Assistant"
  New-Item -ItemType Directory -Path $startMenu -Force | Out-Null
  $shortcutPath = Join-Path $startMenu "Start semantic service.lnk"
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($shortcutPath)
  $shortcut.TargetPath = Join-Path $targetRoot "start-service.cmd"
  $shortcut.WorkingDirectory = $targetRoot
  $shortcut.Description = "Local semantic search service for Outlook AI Assistant"
  $shortcut.Save()
  Write-Host "Start Menu shortcut created:"
  Write-Host "  $shortcutPath"
}
catch {
  Write-Host "Could not create a Start Menu shortcut: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Run:  $(Join-Path $targetRoot 'download-model.cmd')"
Write-Host "  2. Run:  $(Join-Path $targetRoot 'start-service.cmd')"
Write-Host "     (or use the Start Menu shortcut)"
