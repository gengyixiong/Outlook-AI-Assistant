[CmdletBinding()]
param(
  [string]$TargetUserSid = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$connectClsid = "{8A435B2A-4E46-4557-9300-BD94F8C92F82}"
$taskPaneClsid = "{60DF60D1-BB82-4606-B245-F79DD62C99AC}"
$connectProgId = "OutlookAiAssistant.Connect"
$taskPaneProgId = "OutlookAiAssistant.TaskPane"
$sourceDll = Join-Path $PSScriptRoot "OutlookAiAssistant.dll"
$productRoot = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$installRoot = Join-Path $productRoot "app"
$installedDll = Join-Path $installRoot "OutlookAiAssistant.dll"

if (-not [string]::IsNullOrWhiteSpace($TargetUserSid) -and
    $TargetUserSid -notmatch '^S-\d(-\d+)+$') {
  throw "TargetUserSid is not a valid Windows SID."
}

$registryRoot = if ([string]::IsNullOrWhiteSpace($TargetUserSid)) {
  "HKCU:"
}
else {
  "Registry::HKEY_USERS\$TargetUserSid"
}

if (-not (Test-Path -LiteralPath $registryRoot)) {
  throw "The target user registry hive is not loaded: $registryRoot"
}

function Set-RegistryDefaultValue {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]$Value
  )

  New-Item -Path $Path -Force | Out-Null
  Set-Item -Path $Path -Value $Value
}

function Register-ManagedComClass {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Clsid,
    [Parameter(Mandatory = $true)]
    [string]$ProgId,
    [Parameter(Mandatory = $true)]
    [string]$ClassName,
    [Parameter(Mandatory = $true)]
    [string]$AssemblyFullName,
    [Parameter(Mandatory = $true)]
    [string]$AssemblyVersion,
    [Parameter(Mandatory = $true)]
    [string]$CodeBase
  )

  $classRoot = Join-Path $script:registryRoot `
    "Software\Classes\CLSID\$Clsid"
  $inprocRoot = Join-Path $classRoot "InprocServer32"
  $versionRoot = Join-Path $inprocRoot $AssemblyVersion

  Set-RegistryDefaultValue -Path $classRoot -Value $ClassName
  Set-RegistryDefaultValue -Path $inprocRoot -Value "mscoree.dll"
  New-ItemProperty -Path $inprocRoot -Name "ThreadingModel" `
    -Value "Both" -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $inprocRoot -Name "Class" `
    -Value $ClassName -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $inprocRoot -Name "Assembly" `
    -Value $AssemblyFullName -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $inprocRoot -Name "RuntimeVersion" `
    -Value "v4.0.30319" -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $inprocRoot -Name "CodeBase" `
    -Value $CodeBase -PropertyType String -Force | Out-Null

  Set-RegistryDefaultValue -Path $versionRoot -Value "mscoree.dll"
  New-ItemProperty -Path $versionRoot -Name "Class" `
    -Value $ClassName -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $versionRoot -Name "Assembly" `
    -Value $AssemblyFullName -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $versionRoot -Name "RuntimeVersion" `
    -Value "v4.0.30319" -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $versionRoot -Name "CodeBase" `
    -Value $CodeBase -PropertyType String -Force | Out-Null

  Set-RegistryDefaultValue -Path (Join-Path $classRoot "ProgId") `
    -Value $ProgId
  Set-RegistryDefaultValue `
    -Path (Join-Path $classRoot "VersionIndependentProgID") `
    -Value $ProgId
  Set-RegistryDefaultValue `
    -Path (Join-Path $classRoot `
      "Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}") `
    -Value ""

  Set-RegistryDefaultValue `
    -Path (Join-Path $script:registryRoot "Software\Classes\$ProgId") `
    -Value $ClassName
  Set-RegistryDefaultValue `
    -Path (Join-Path $script:registryRoot "Software\Classes\$ProgId\CLSID") `
    -Value $Clsid
}

if (-not (Test-Path -LiteralPath $sourceDll -PathType Leaf)) {
  throw "OutlookAiAssistant.dll is missing from the installer folder."
}

if (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue) {
  throw "Please close classic Outlook before installing or updating the add-in."
}

New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
Copy-Item -LiteralPath $sourceDll -Destination $installedDll -Force

$assemblyName = [System.Reflection.AssemblyName]::GetAssemblyName($installedDll)
$assemblyFullName = $assemblyName.FullName
$assemblyVersion = $assemblyName.Version.ToString()
$codeBase = (New-Object System.Uri($installedDll)).AbsoluteUri

Register-ManagedComClass -Clsid $connectClsid `
  -ProgId $connectProgId `
  -ClassName "OutlookAiAssistant.AddIn.Connect" `
  -AssemblyFullName $assemblyFullName `
  -AssemblyVersion $assemblyVersion `
  -CodeBase $codeBase
Register-ManagedComClass -Clsid $taskPaneClsid `
  -ProgId $taskPaneProgId `
  -ClassName "OutlookAiAssistant.UI.AssistantPaneControl" `
  -AssemblyFullName $assemblyFullName `
  -AssemblyVersion $assemblyVersion `
  -CodeBase $codeBase

$addinKey = Join-Path $registryRoot `
  "Software\Microsoft\Office\Outlook\Addins\$connectProgId"
New-Item -Path $addinKey -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "FriendlyName" `
  -Value "Outlook AI Assistant" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "Description" `
  -Value "Manual email summaries and natural-language local search" `
  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "LoadBehavior" `
  -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "CommandLineSafe" `
  -Value 0 -PropertyType DWord -Force | Out-Null

Write-Host ""
Write-Host "Outlook AI Assistant was installed for the current Windows user."
Write-Host "Start classic Outlook. The AI assistant pane should appear on the right."
Write-Host "No administrator permission was used."
