[CmdletBinding()]
param(
  [ValidateSet("Debug", "Release")]
  [string]$Configuration = "Release",
  [switch]$SkipTests,
  [switch]$SkipPackage
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$desktopRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $desktopRoot "src\OutlookAiAssistant"
$testRoot = Join-Path $desktopRoot "tests\OutlookAiAssistant.UnitTests"
$buildRoot = Join-Path $desktopRoot ("build\" + $Configuration)
$releaseRoot = Join-Path $desktopRoot "release"
$packageRoot = Join-Path $releaseRoot "package"

function Find-FirstFile {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Candidates,
    [Parameter(Mandatory = $true)]
    [string]$Description
  )

  foreach ($candidate in $Candidates) {
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
      return (Resolve-Path -LiteralPath $candidate).Path
    }
  }

  throw "Cannot find $Description. Install classic Outlook and .NET Framework 4.x."
}

function Find-GacAssembly {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Roots,
    [Parameter(Mandatory = $true)]
    [string]$FileName,
    [Parameter(Mandatory = $true)]
    [string]$Description
  )

  foreach ($root in $Roots) {
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
      continue
    }

    $match = Get-ChildItem -LiteralPath $root -Recurse -Filter $FileName `
      -ErrorAction SilentlyContinue |
      Sort-Object FullName -Descending |
      Select-Object -First 1
    if ($null -ne $match) {
      return $match.FullName
    }
  }

  throw "Cannot find $Description ($FileName). Repair Microsoft Office."
}

$compiler = Find-FirstFile -Description ".NET Framework C# compiler" -Candidates @(
  "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
  "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$frameworkDirectory = Split-Path -Parent $compiler
$gacRoots = @(
  "$env:WINDIR\assembly\GAC_MSIL",
  "$env:WINDIR\assembly\GAC_32",
  "$env:WINDIR\assembly\GAC"
)
$officeInterop = Find-GacAssembly -Roots $gacRoots `
  -FileName "OFFICE.DLL" -Description "Microsoft Office PIA"
$outlookInterop = Find-GacAssembly -Roots $gacRoots `
  -FileName "Microsoft.Office.Interop.Outlook.dll" `
  -Description "Microsoft Outlook PIA"
$extensibility = Find-GacAssembly -Roots $gacRoots `
  -FileName "Extensibility.dll" -Description "Office Extensibility PIA"

New-Item -ItemType Directory -Path $buildRoot -Force | Out-Null
$assemblyPath = Join-Path $buildRoot "OutlookAiAssistant.dll"
$pdbPath = Join-Path $buildRoot "OutlookAiAssistant.pdb"
$sourceFiles = Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter "*.cs" |
  Sort-Object FullName |
  ForEach-Object FullName

if ($sourceFiles.Count -eq 0) {
  throw "No C# source files were found."
}

$references = @(
  (Join-Path $frameworkDirectory "System.dll"),
  (Join-Path $frameworkDirectory "System.Core.dll"),
  (Join-Path $frameworkDirectory "System.Drawing.dll"),
  (Join-Path $frameworkDirectory "System.Security.dll"),
  (Join-Path $frameworkDirectory "System.Web.Extensions.dll"),
  (Join-Path $frameworkDirectory "System.Windows.Forms.dll"),
  $officeInterop,
  $outlookInterop,
  $extensibility
)

$compilerArguments = @(
  "/target:library",
  "/platform:anycpu",
  "/codepage:65001",
  "/warn:4",
  "/out:$assemblyPath",
  "/pdb:$pdbPath"
)

if ($Configuration -eq "Release") {
  $compilerArguments += @("/optimize+", "/debug:pdbonly")
}
else {
  $compilerArguments += @("/optimize-", "/debug:full", "/define:DEBUG;TRACE")
}

foreach ($reference in $references) {
  $compilerArguments += "/reference:$reference"
}
$compilerArguments += $sourceFiles

Write-Host "Building OutlookAiAssistant.dll ($Configuration)..."
& $compiler @compilerArguments
if ($LASTEXITCODE -ne 0) {
  throw "Add-in compilation failed with exit code $LASTEXITCODE."
}

if (-not $SkipTests) {
  $testExecutable = Join-Path $buildRoot "OutlookAiAssistant.UnitTests.exe"
  $testSources = Get-ChildItem -LiteralPath $testRoot -Recurse -Filter "*.cs" |
    Sort-Object FullName |
    ForEach-Object FullName
  $testArguments = @(
    "/target:exe",
    "/platform:anycpu",
    "/codepage:65001",
    "/warn:4",
    "/out:$testExecutable",
    "/reference:$assemblyPath",
    ("/reference:" + (Join-Path $frameworkDirectory "System.dll")),
    ("/reference:" + (Join-Path $frameworkDirectory "System.Core.dll")),
    ("/reference:" + (Join-Path $frameworkDirectory "System.Web.Extensions.dll")),
    ("/reference:" + (Join-Path $frameworkDirectory "System.Drawing.dll")),
    ("/reference:" + (Join-Path $frameworkDirectory "System.Security.dll")),
    ("/reference:" + (Join-Path $frameworkDirectory "System.Windows.Forms.dll"))
  )
  $testArguments += $testSources

  Write-Host "Building unit tests..."
  & $compiler @testArguments
  if ($LASTEXITCODE -ne 0) {
    throw "Unit test compilation failed with exit code $LASTEXITCODE."
  }

  Write-Host "Running unit tests..."
  & $testExecutable
  if ($LASTEXITCODE -ne 0) {
    throw "Unit tests failed with exit code $LASTEXITCODE."
  }
}

if (-not $SkipPackage) {
  $resolvedDesktopRoot = (Resolve-Path -LiteralPath $desktopRoot).Path
  $resolvedReleaseRoot = [System.IO.Path]::GetFullPath($releaseRoot)
  if (-not $resolvedReleaseRoot.StartsWith(
      $resolvedDesktopRoot,
      [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to clean a release path outside the project: $resolvedReleaseRoot"
  }

  if (Test-Path -LiteralPath $packageRoot) {
    Remove-Item -LiteralPath $packageRoot -Recurse -Force
  }

  New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
  Copy-Item -LiteralPath $assemblyPath -Destination $packageRoot
  if (Test-Path -LiteralPath $pdbPath) {
    Copy-Item -LiteralPath $pdbPath -Destination $packageRoot
  }

  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\install.ps1") `
    -Destination $packageRoot
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\uninstall.ps1") `
    -Destination $packageRoot
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\Install.cmd") `
    -Destination $packageRoot
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\Uninstall.cmd") `
    -Destination $packageRoot
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\README.txt") `
    -Destination $packageRoot
  $packageDocs = Join-Path $packageRoot "Documentation"
  New-Item -ItemType Directory -Path $packageDocs -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $desktopRoot "..\docs\PRIVACY.md") `
    -Destination $packageDocs
  Copy-Item -LiteralPath (Join-Path $desktopRoot "..\docs\INSTALLATION.md") `
    -Destination $packageDocs

  $zipPath = Join-Path $releaseRoot "Outlook-AI-Assistant-v0.2.1.zip"
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  Compress-Archive -Path (Join-Path $packageRoot "*") `
    -DestinationPath $zipPath -CompressionLevel Optimal
  Write-Host "Package created: $zipPath"
}

Write-Host "Build completed successfully."
