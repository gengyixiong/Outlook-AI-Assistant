$ErrorActionPreference = 'Stop'

$projectRootPath = Split-Path -Parent $PSScriptRoot
$serverScriptPath = Join-Path $PSScriptRoot 'serve-dist.mjs'
$globalNodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
$explicitNodePath = $env:OUTLOOK_AI_NODE_PATH

if ($explicitNodePath -and (Test-Path -LiteralPath $explicitNodePath)) {
    $nodeExecutablePath = $explicitNodePath
}
elseif ($globalNodeCommand) {
    $nodeExecutablePath = $globalNodeCommand.Source
}
else {
    $userProfilePath = [Environment]::GetFolderPath('UserProfile')
    $codexNodePath = Join-Path $userProfilePath '.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'

    if (-not (Test-Path -LiteralPath $codexNodePath)) {
        throw 'Node.js was not found. Install Node.js 20 or later, then run this script again.'
    }

    $nodeExecutablePath = $codexNodePath
}

Write-Host "Project: $projectRootPath"
Write-Host "Node.js: $nodeExecutablePath"
Write-Host ''

& $nodeExecutablePath $serverScriptPath
exit $LASTEXITCODE
