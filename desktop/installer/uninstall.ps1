[CmdletBinding()]
param(
  [switch]$RemoveSettings
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$connectClsid = "{8A435B2A-4E46-4557-9300-BD94F8C92F82}"
$taskPaneClsid = "{60DF60D1-BB82-4606-B245-F79DD62C99AC}"
$connectProgId = "OutlookAiAssistant.Connect"
$taskPaneProgId = "OutlookAiAssistant.TaskPane"
$productRoot = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$installRoot = Join-Path $productRoot "app"

if (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue) {
  throw "Please close classic Outlook before uninstalling the add-in."
}

$registryPaths = @(
  "HKCU:\Software\Microsoft\Office\Outlook\Addins\$connectProgId",
  "HKCU:\Software\Classes\$connectProgId",
  "HKCU:\Software\Classes\$taskPaneProgId",
  "HKCU:\Software\Classes\CLSID\$connectClsid",
  "HKCU:\Software\Classes\CLSID\$taskPaneClsid"
)

foreach ($path in $registryPaths) {
  if (Test-Path -LiteralPath $path) {
    Remove-Item -LiteralPath $path -Recurse -Force
  }
}

$resolvedProductRoot = [System.IO.Path]::GetFullPath($productRoot)
$resolvedInstallRoot = [System.IO.Path]::GetFullPath($installRoot)
if (-not $resolvedInstallRoot.StartsWith(
    $resolvedProductRoot,
    [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "Refusing to delete an unexpected install path: $resolvedInstallRoot"
}

if (Test-Path -LiteralPath $installRoot) {
  Remove-Item -LiteralPath $installRoot -Recurse -Force
}

if ($RemoveSettings -and (Test-Path -LiteralPath $productRoot)) {
  Remove-Item -LiteralPath $productRoot -Recurse -Force
  Write-Host "The add-in, encrypted API key, settings and logs were removed."
}
else {
  Write-Host "The add-in was removed. Settings and encrypted API key were preserved."
}

