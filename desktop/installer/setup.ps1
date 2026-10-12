# =====================================================================
#  Outlook AI Assistant - unified installer  (setup.ps1)
#
#  One entry point for the whole installation. Safe to run repeatedly:
#  it detects what is already installed and only does what is needed.
#
#  It performs, in order:
#    1. Check Python 3.10-3.12
#    2. Install / update the Outlook add-in DLL (COM registration)
#    3. Install / update the semantic search service code
#    4. Create the Python venv and dependencies (only once)
#    5. Download the embedding model (only once)
#
#  Data is never destroyed: the venv, the model cache, the semantic index,
#  settings and the encrypted API key are always preserved.
#
#  Indexing (reading mail) is deliberately NOT done here. It is long
#  running and needs Outlook running; do it from the add-in pane.
#
#  Usage:
#    setup.cmd                      # normal install / update
#    setup.cmd -CheckOnly           # report state, change nothing
#    setup.cmd -Force               # redo every step even if up to date
#    setup.cmd -SkipModel           # do not download the model
#    setup.cmd -NoPause             # do not wait for a keypress
# =====================================================================

[CmdletBinding()]
param(
  [switch]$CheckOnly,
  [switch]$Force,
  [switch]$SkipModel,
  [switch]$NoPause
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$sourceDir = $PSScriptRoot
$productRoot = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$appDir = Join-Path $productRoot "app"
$serviceDir = Join-Path $productRoot "embedding_service"
$modelDir = Join-Path $productRoot "models"
$versionMarker = Join-Path $productRoot ".installed-version"

$script:Problems = New-Object System.Collections.Generic.List[string]
$script:Changed = $false

# ---------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------
function Write-Step {
  param([string]$Text)
  Write-Host ""
  Write-Host ("=" * 68) -ForegroundColor DarkGray
  Write-Host $Text -ForegroundColor Cyan
  Write-Host ("=" * 68) -ForegroundColor DarkGray
}

function Write-Ok      { param([string]$T) Write-Host "  [OK]   $T" -ForegroundColor Green }
function Write-Skip    { param([string]$T) Write-Host "  [跳过] $T" -ForegroundColor DarkGray }
function Write-Action  { param([string]$T) Write-Host "  [执行] $T" -ForegroundColor Yellow }
function Write-WarnMsg { param([string]$T) Write-Host "  [注意] $T" -ForegroundColor Yellow }
function Write-ErrMsg  { param([string]$T) Write-Host "  [错误] $T" -ForegroundColor Red }

function Get-PackageVersion {
  $file = Join-Path $sourceDir "VERSION"
  if (-not (Test-Path -LiteralPath $file)) {
    throw "VERSION file missing from the package. Re-extract the release ZIP."
  }
  return (Get-Content -LiteralPath $file -Raw).Trim()
}

function Get-InstalledVersion {
  if (Test-Path -LiteralPath $versionMarker) {
    return (Get-Content -LiteralPath $versionMarker -Raw).Trim()
  }
  return ""
}

function Get-AddInDllVersion {
  $dll = Join-Path $appDir "OutlookAiAssistant.dll"
  if (-not (Test-Path -LiteralPath $dll)) { return "" }
  try {
    return (Get-Item -LiteralPath $dll).VersionInfo.FileVersion
  }
  catch { return "" }
}

function Test-ServiceInstalled {
  return (Test-Path -LiteralPath (Join-Path $serviceDir "app\main.py"))
}

function Test-VenvReady {
  return (Test-Path -LiteralPath `
    (Join-Path $serviceDir ".venv\Scripts\python.exe"))
}

function Test-VenvNeedsCudaUpgrade {
  $pythonExe = Join-Path $serviceDir ".venv\Scripts\python.exe"
  if (-not (Test-Path -LiteralPath $pythonExe)) { return $false }

  $hasNvidia = $false
  $nvSmi = $null
  $candidates = @(
    "$env:SystemRoot\System32\nvidia-smi.exe",
    "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
    "${env:ProgramW6432}\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
  )
  foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) { $nvSmi = $c; break }
  }
  if (-not $nvSmi) {
    $where = (Get-Command nvidia-smi -ErrorAction SilentlyContinue)
    if ($where) { $nvSmi = $where.Source }
  }
  if ($nvSmi) {
    & $nvSmi -L >$null 2>&1
    if ($LASTEXITCODE -eq 0) { $hasNvidia = $true }
  }

  if (-not $hasNvidia) { return $false }

  try {
    $out = & $pythonExe -c "import torch; print(torch.cuda.is_available())" 2>$null
    if ($out -notmatch "True") {
      return $true
    }
  }
  catch {
    return $true
  }

  return $false
}

function Test-ModelReady {
  # Mirror the service check: a hub cache folder holding the weights, either
  # in our own model dir or in the legacy Hugging Face location.
  $safe = "models--google--embeddinggemma-2"
  $roots = @(
    (Join-Path $modelDir "huggingface\hub"),
    (Join-Path $modelDir "hub"),
    (Join-Path $env:USERPROFILE ".cache\huggingface\hub")
  )
  foreach ($root in $roots) {
    $folder = Join-Path $root $safe
    if (-not (Test-Path -LiteralPath $folder)) { continue }
    $hits = Get-ChildItem -LiteralPath $folder -Recurse -Filter `
      "model.safetensors" -ErrorAction SilentlyContinue
    if ($hits) { return $true }
  }
  return $false
}

function Find-SupportedPython {
  # Returns a hashtable with Path and Version, or $null.
  $launcher = Get-Command py -ErrorAction SilentlyContinue
  if ($launcher) {
    foreach ($v in @("3.12", "3.11", "3.10")) {
      try {
        $exe = & py "-$v" -c "import sys;print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $exe) {
          return @{ Path = $exe.Trim(); Version = $v }
        }
      }
      catch { }
    }
  }

  $candidates = @()
  $onPath = Get-Command python -ErrorAction SilentlyContinue
  if ($onPath -and $onPath.Source -and
      $onPath.Source -notmatch 'WindowsApps') {
    $candidates += $onPath.Source
  }
  foreach ($dir in @(
      (Join-Path $env:LOCALAPPDATA "Programs\Python\Python312"),
      (Join-Path $env:LOCALAPPDATA "Programs\Python\Python311"),
      (Join-Path $env:LOCALAPPDATA "Programs\Python\Python310"),
      (Join-Path $env:ProgramFiles "Python312"),
      (Join-Path $env:ProgramFiles "Python311"),
      (Join-Path $env:ProgramFiles "Python310"))) {
    $exe = Join-Path $dir "python.exe"
    if (Test-Path -LiteralPath $exe) { $candidates += $exe }
  }

  foreach ($exe in $candidates) {
    try {
      $out = & $exe -c `
        "import sys;print('%d.%d' % sys.version_info[:2])" 2>$null
      if ($LASTEXITCODE -ne 0) { continue }
      $ver = $out.Trim()
      if ($ver -in @("3.10", "3.11", "3.12")) {
        return @{ Path = $exe; Version = $ver }
      }
    }
    catch { }
  }

  return $null
}

# ---------------------------------------------------------------------
# banner
# ---------------------------------------------------------------------
Write-Host ""
Write-Host "Outlook AI Assistant - 安装/更新" -ForegroundColor White
Write-Host ("=" * 68) -ForegroundColor DarkGray

$packageVersion = Get-PackageVersion
$installedVersion = Get-InstalledVersion

Write-Host ("打包版本   : {0}" -f $packageVersion)
if ($installedVersion) {
  Write-Host ("已安装版本 : {0}" -f $installedVersion)
}
else {
  Write-Host "已安装版本 : （未安装）"
}
if ($CheckOnly) {
  Write-Host ""
  Write-Host "检查模式：不会修改任何文件。" -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
# step 1: Python
# ---------------------------------------------------------------------
Write-Step "1/5  检查 Python 环境"

$python = Find-SupportedPython
if (-not $python) {
  Write-ErrMsg "未找到受支持的 Python (3.10 / 3.11 / 3.12)。"
  Write-Host ""
  Write-Host "  请安装 Python 3.12："
  Write-Host "    https://www.python.org/downloads/"
  Write-Host "  安装时勾选 'Add python.exe to PATH'。"
  Write-Host ""
  Write-Host "  注意：Python 3.13 及以上不可用，PyTorch 没有对应的预编译包。"
  Write-Host "  如果 python 命令会打开应用商店，请关闭以下别名："
  Write-Host "    设置 > 应用 > 高级应用设置 > 应用执行别名"
  Write-Host "    关闭 python.exe 与 python3.exe"
  $script:Problems.Add("缺少受支持的 Python")
}
else {
  Write-Ok ("Python {0} -> {1}" -f $python.Version, $python.Path)
}

# ---------------------------------------------------------------------
# step 2: add-in DLL
# ---------------------------------------------------------------------
Write-Step "2/5  安装 / 更新 Outlook 加载项"

$dllVersion = Get-AddInDllVersion
$needAddIn = $Force -or ($dllVersion -ne "$packageVersion.0")

if ($CheckOnly) {
  if ($dllVersion -eq "$packageVersion.0") {
    Write-Ok ("加载项已是最新 ({0})" -f $dllVersion)
  }
  elseif ($dllVersion) {
    Write-WarnMsg ("加载项版本 {0}，包内为 {1} —— 需要更新" -f `
      $dllVersion, "$packageVersion.0")
  }
  else {
    Write-WarnMsg "加载项未安装"
  }
}
elseif (-not $needAddIn) {
  Write-Skip ("加载项已是最新 ({0})" -f $dllVersion)
}
else {
  if (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue) {
    Write-ErrMsg "请先完全退出 Outlook，再重新运行本脚本。"
    $script:Problems.Add("Outlook 正在运行，无法更新加载项")
  }
  else {
    Write-Action "注册并复制加载项 DLL"
    $installer = Join-Path $sourceDir "install.ps1"
    if (-not (Test-Path -LiteralPath $installer)) {
      Write-ErrMsg "包内缺少 install.ps1，请重新解压完整发布包。"
      $script:Problems.Add("install.ps1 缺失")
    }
    else {
      & powershell.exe -NoProfile -ExecutionPolicy Bypass `
        -File $installer
      if ($LASTEXITCODE -ne 0) {
        Write-ErrMsg ("加载项安装失败，退出码 {0}" -f $LASTEXITCODE)
        $script:Problems.Add("加载项安装失败")
      }
      else {
        $script:Changed = $true
        Write-Ok ("加载项已更新到 {0}" -f (Get-AddInDllVersion))
      }
    }
  }
}

# ---------------------------------------------------------------------
# step 3: service code
# ---------------------------------------------------------------------
Write-Step "3/5  安装 / 更新语义搜索服务代码"

$needService = $Force -or (-not (Test-ServiceInstalled)) `
  -or ($installedVersion -ne $packageVersion)

if ($CheckOnly) {
  if (Test-ServiceInstalled) {
    Write-Ok ("服务代码已安装: {0}" -f $serviceDir)
  }
  else {
    Write-WarnMsg "服务代码未安装"
  }
}
elseif (-not $needService) {
  Write-Skip "服务代码已是最新"
}
else {
  Write-Action "复制服务代码（保留 .venv 与模型缓存）"
  $serviceInstaller = Join-Path $sourceDir "embedding_service\install-service.ps1"
  if (-not (Test-Path -LiteralPath $serviceInstaller)) {
    Write-ErrMsg "包内缺少 embedding_service\install-service.ps1。"
    $script:Problems.Add("服务安装脚本缺失")
  }
  else {
    # -SkipVenv: step 4 handles the venv so the progress stays clear.
    & powershell.exe -NoProfile -ExecutionPolicy Bypass `
      -File $serviceInstaller -SkipVenv
    if ($LASTEXITCODE -ne 0) {
      Write-ErrMsg ("服务代码安装失败，退出码 {0}" -f $LASTEXITCODE)
      $script:Problems.Add("服务代码安装失败")
    }
    else {
      $script:Changed = $true
      Write-Ok "服务代码已更新"
    }
  }
}

# ---------------------------------------------------------------------
# step 4: venv
# ---------------------------------------------------------------------
Write-Step "4/5  Python 依赖环境 (venv)"

$cudaUpgradeNeeded = Test-VenvNeedsCudaUpgrade
$needVenv = $Force -or (-not (Test-VenvReady)) -or $cudaUpgradeNeeded

if ($CheckOnly) {
  if (Test-VenvReady) {
    if ($cudaUpgradeNeeded) {
      Write-WarnMsg "依赖环境已就绪但未启用 CUDA (检测到 NVIDIA 显卡，需更新为 cu130)"
    } else {
      Write-Ok "依赖环境已就绪"
    }
  }
  else { Write-WarnMsg "依赖环境未创建" }
}
elseif (-not $needVenv) {
  Write-Skip "依赖环境已就绪（不会重复安装）"
}
elseif (-not $python) {
  Write-ErrMsg "跳过：没有可用的 Python。"
}
elseif (-not (Test-ServiceInstalled)) {
  Write-ErrMsg "跳过：服务代码尚未安装。"
}
else {
  if ($cudaUpgradeNeeded) {
    Write-Action "检测到 NVIDIA 显卡但现有环境未启用 CUDA，正在更新为 CUDA (cu130) 环境..."
  } else {
    Write-Action "创建虚拟环境并安装依赖（首次较慢，含 PyTorch）"
  }
  $setup = Join-Path $serviceDir "setup-venv.cmd"
  & cmd.exe /c "`"$setup`" --no-pause"
  if ($LASTEXITCODE -ne 0) {
    Write-ErrMsg ("依赖安装失败，退出码 {0}" -f $LASTEXITCODE)
    $script:Problems.Add("依赖安装失败")
  }
  else {
    $script:Changed = $true
    Write-Ok "依赖环境已创建 / 更新"
  }
}

# ---------------------------------------------------------------------
# step 5: model
# ---------------------------------------------------------------------
Write-Step "5/5  本地语义模型"

$needModel = $Force -or (-not (Test-ModelReady))

if ($SkipModel) {
  Write-Skip "按要求跳过模型下载 (-SkipModel)"
}
elseif ($CheckOnly) {
  if (Test-ModelReady) { Write-Ok "模型已缓存" }
  else { Write-WarnMsg "模型未下载（约 1.5 GB，需要联网）" }
}
elseif (-not $needModel) {
  Write-Skip "模型已缓存（不会重复下载）"
}
elseif (-not (Test-VenvReady)) {
  Write-ErrMsg "跳过：依赖环境尚未就绪。"
}
else {
  Write-Action "下载模型（约 1.5 GB，需要联网；之后可离线使用）"
  $download = Join-Path $serviceDir "download-model.cmd"
  & cmd.exe /c "`"$download`" --no-pause"
  if ($LASTEXITCODE -ne 0) {
    Write-ErrMsg ("模型下载失败，退出码 {0}" -f $LASTEXITCODE)
    $script:Problems.Add("模型下载失败")
  }
  else {
    $script:Changed = $true
    Write-Ok "模型已就绪"
  }
}

# ---------------------------------------------------------------------
# record version
# ---------------------------------------------------------------------
if (-not $CheckOnly -and $script:Problems.Count -eq 0) {
  try {
    New-Item -ItemType Directory -Path $productRoot -Force | Out-Null
    Set-Content -LiteralPath $versionMarker -Value $packageVersion `
      -Encoding ASCII
  }
  catch {
    Write-WarnMsg ("无法写入版本标记: {0}" -f $_.Exception.Message)
  }
}

# ---------------------------------------------------------------------
# summary
# ---------------------------------------------------------------------
Write-Step "结果"

if ($CheckOnly) {
  Write-Host "  检查完成（未做任何修改）。" -ForegroundColor Green
  Write-Host ""
  Write-Host "  直接运行 setup.cmd（不带 -CheckOnly）即可补齐上面标记为"
  Write-Host "  [注意] 的项目。已完成的步骤会自动跳过。"
}
elseif ($script:Problems.Count -eq 0) {
  Write-Host "  所有步骤已完成。" -ForegroundColor Green
  Write-Host ""
  Write-Host "  下一步（在 Outlook 中）："
  Write-Host "    1. 打开 Outlook，右侧 'AI 邮件助手' -> '语义搜索'"
  Write-Host "    2. 点击 '启动服务'（后台运行，无窗口）"
  Write-Host "    3. 点击 '建立 / 更新语义索引'（首次较慢）"
  Write-Host "    4. 输入中文描述开始搜索，双击结果卡片打开原邮件"
  Write-Host ""
  Write-Host "  解压出来的这个文件夹现在可以删除。"
  Write-Host "  如需卸载服务：对 uninstall-service.ps1 运行 PowerShell。"
}
else {
  Write-Host "  以下步骤未成功：" -ForegroundColor Red
  foreach ($p in $script:Problems) { Write-Host ("    - " + $p) -ForegroundColor Red }
  Write-Host ""
  Write-Host "  可运行 diagnose.cmd 生成诊断报告。" -ForegroundColor Yellow
}

Write-Host ""
Write-Host ("数据目录: {0}" -f $productRoot) -ForegroundColor DarkGray
Write-Host "升级时会保留：依赖环境、模型缓存、语义索引、设置与 API Key。" -ForegroundColor DarkGray

if (-not $NoPause) {
  Write-Host ""
  Write-Host "按任意键关闭..." -ForegroundColor Cyan
  try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { }
}

if ($script:Problems.Count -eq 0) { exit 0 } else { exit 1 }
