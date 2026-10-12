# Removes the installed local semantic search service.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall-service.ps1
#
# Options:
#   -RemoveModel   also delete the downloaded embedding model cache
#   -RemoveIndex   also delete the local semantic index (mail.db and vectors)
#
# By default the model cache and the semantic index are PRESERVED so a
# reinstall does not have to re-download ~1.5 GB or rebuild the index.

[CmdletBinding()]
param(
  [switch]$RemoveModel,
  [switch]$RemoveIndex
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$productRoot = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$targetRoot = Join-Path $productRoot "embedding_service"

if (Test-Path -LiteralPath $targetRoot) {
  Remove-Item -LiteralPath $targetRoot -Recurse -Force
  Write-Host "Removed: $targetRoot"
}
else {
  Write-Host "Nothing to remove at: $targetRoot"
}

# Start Menu shortcut
$shortcutFolder = Join-Path $env:APPDATA `
  "Microsoft\Windows\Start Menu\Programs\Outlook AI Assistant"
$shortcutPath = Join-Path $shortcutFolder "Start semantic service.lnk"
if (Test-Path -LiteralPath $shortcutPath) {
  Remove-Item -LiteralPath $shortcutPath -Force
  Write-Host "Removed Start Menu shortcut."
}
if (Test-Path -LiteralPath $shortcutFolder) {
  $remaining = Get-ChildItem -LiteralPath $shortcutFolder -Force `
    -ErrorAction SilentlyContinue
  if ($null -eq $remaining -or $remaining.Count -eq 0) {
    Remove-Item -LiteralPath $shortcutFolder -Force
  }
}

if ($RemoveModel) {
  $modelRoot = Join-Path $productRoot "models"
  if (Test-Path -LiteralPath $modelRoot) {
    Remove-Item -LiteralPath $modelRoot -Recurse -Force
    Write-Host "Removed the downloaded model cache: $modelRoot"
  }
}
else {
  Write-Host "Model cache preserved (use -RemoveModel to delete it)."
}

if ($RemoveIndex) {
  $indexRoot = Join-Path $productRoot "semantic-index"
  if (Test-Path -LiteralPath $indexRoot) {
    Remove-Item -LiteralPath $indexRoot -Recurse -Force
    Write-Host "Removed the local semantic index: $indexRoot"
  }
}
else {
  Write-Host "Semantic index preserved (use -RemoveIndex to delete it)."
}

Write-Host ""
Write-Host "The semantic service was removed. Outlook itself is unaffected."
