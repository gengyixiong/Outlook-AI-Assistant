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
$repoRoot = Split-Path -Parent $desktopRoot
$sourceRoot = Join-Path $desktopRoot "src\OutlookAiAssistant"
$testRoot = Join-Path $desktopRoot "tests\OutlookAiAssistant.UnitTests"
$buildRoot = Join-Path $desktopRoot ("build\" + $Configuration)
$releaseRoot = Join-Path $desktopRoot "release"
$packageRoot = Join-Path $releaseRoot "package"

# Single source of truth for the product version.
$versionFile = Join-Path $repoRoot "VERSION"
if (-not (Test-Path -LiteralPath $versionFile)) {
  throw "VERSION file is missing at $versionFile"
}
$productVersion = (Get-Content -LiteralPath $versionFile -Raw).Trim()
if ($productVersion -notmatch '^\d+\.\d+\.\d+$') {
  throw "VERSION must be MAJOR.MINOR.PATCH, got '$productVersion'"
}
$assemblyVersion = "$productVersion.0"
Write-Host "Product version: $productVersion"

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

  throw "Cannot find $Description. Install the .NET Framework 4.8 Developer Pack."
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
$frameworkReferenceCandidates = @()
if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
  $frameworkReferenceCandidates += Join-Path ${env:ProgramFiles(x86)} `
    "Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\mscorlib.dll"
}
if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
  $frameworkReferenceCandidates += Join-Path $env:ProgramFiles `
    "Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\mscorlib.dll"
}
$frameworkReferenceAssembly = Find-FirstFile `
  -Description ".NET Framework 4.8 reference assemblies" `
  -Candidates $frameworkReferenceCandidates
$frameworkReferenceDirectory = Split-Path -Parent $frameworkReferenceAssembly
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

# Generate the assembly version file from VERSION so the DLL version and the
# package name can never drift apart.
$versionSource = Join-Path $buildRoot "GeneratedVersion.cs"
@"
using System.Reflection;
[assembly: AssemblyVersion("$assemblyVersion")]
[assembly: AssemblyFileVersion("$assemblyVersion")]
"@ | Set-Content -LiteralPath $versionSource -Encoding UTF8

$sourceFiles = Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter "*.cs" |
  Sort-Object FullName |
  ForEach-Object FullName
$sourceFiles += $versionSource

if ($sourceFiles.Count -eq 0) {
  throw "No C# source files were found."
}

$frameworkReferences = @(
  $frameworkReferenceAssembly,
  (Join-Path $frameworkReferenceDirectory "System.dll"),
  (Join-Path $frameworkReferenceDirectory "System.Core.dll"),
  (Join-Path $frameworkReferenceDirectory "System.Drawing.dll"),
  (Join-Path $frameworkReferenceDirectory "System.Security.dll"),
  (Join-Path $frameworkReferenceDirectory "System.Web.Extensions.dll"),
  (Join-Path $frameworkReferenceDirectory "System.Windows.Forms.dll")
)
$references = $frameworkReferences + @(
  $officeInterop,
  $outlookInterop,
  $extensibility
)

$compilerArguments = @(
  "/noconfig",
  "/target:library",
  "/nostdlib+",
  "/platform:anycpu",
  "/codepage:65001",
  "/warn:4",
  "/out:$assemblyPath"
)

if ($Configuration -eq "Release") {
  $compilerArguments += "/optimize+"
}
else {
  $compilerArguments += @(
    "/optimize-",
    "/debug:full",
    "/pdb:$pdbPath",
    "/define:DEBUG;TRACE"
  )
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
    "/noconfig",
    "/target:exe",
    "/nostdlib+",
    "/platform:anycpu",
    "/codepage:65001",
    "/warn:4",
    "/out:$testExecutable",
    "/reference:$assemblyPath"
  )
  foreach ($reference in $frameworkReferences) {
    $testArguments += "/reference:$reference"
  }
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

  # Stamp the product version into the package so setup.cmd can compare
  # against what is already installed.
  Set-Content -LiteralPath (Join-Path $packageRoot "VERSION") `
    -Value $productVersion -Encoding ASCII

  # Ship the unified installer as the single entry point.
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\setup.cmd") `
    -Destination $packageRoot
  Copy-Item -LiteralPath (Join-Path $desktopRoot "installer\setup.ps1") `
    -Destination $packageRoot

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
  Copy-Item -LiteralPath (Join-Path $desktopRoot "..\docs\SEMANTIC-SEARCH.md") `
    -Destination $packageDocs

  # Ship the local semantic search service alongside the add-in. The .venv
  # and the model cache are intentionally excluded: they are large and are
  # created per machine by setup-venv.cmd / download-model.cmd.
  $serviceSource = Join-Path $desktopRoot "..\embedding_service"
  $serviceTarget = Join-Path $packageRoot "embedding_service"
  New-Item -ItemType Directory -Path $serviceTarget -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $serviceSource "README.md") `
    -Destination $serviceTarget
  Copy-Item -LiteralPath (Join-Path $serviceSource "requirements.txt") `
    -Destination $serviceTarget
  foreach ($script in @(
      "setup-venv.cmd", "start-service.cmd", "download-model.cmd",
      "install-service.cmd", "install-service.ps1",
      "uninstall-service.ps1", "diagnose.ps1", "diagnose.cmd")) {
    Copy-Item -LiteralPath (Join-Path $serviceSource $script) `
      -Destination $serviceTarget
  }
  Copy-Item -LiteralPath (Join-Path $serviceSource "app") `
    -Destination $serviceTarget -Recurse
  # Strip any stray caches the copy may have picked up.
  Get-ChildItem -LiteralPath $serviceTarget -Recurse -Directory `
    -Filter "__pycache__" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }

  $zipPath = Join-Path $releaseRoot "Outlook-AI-Assistant-v$productVersion.zip"
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }
  Compress-Archive -Path (Join-Path $packageRoot "*") `
    -DestinationPath $zipPath -CompressionLevel Optimal
  Write-Host "Package created: $zipPath"
}

Write-Host "Build completed successfully."
