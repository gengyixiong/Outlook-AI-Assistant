# =====================================================================
#  Outlook AI Assistant - 诊断脚本 (diagnose.ps1)
#
#  用途：一次性收集排查"启动服务无效 / 模型未就绪 / /health 打不开"
#        所需的全部信息，并把结果写入一个文本文件。
#
#  用法（二选一）：
#    1) 右键本文件 -> "使用 PowerShell 运行"
#    2) 在 PowerShell 中执行：
#         powershell -NoProfile -ExecutionPolicy Bypass -File .\diagnose.ps1
#
#  运行结束后会在桌面生成 OutlookAiAssistant-Diagnose.txt，
#  请把该文件的全部内容发回来。
# =====================================================================

[CmdletBinding()]
param(
  [string]$OutputPath = "",
  [switch]$NoPause
)

$ErrorActionPreference = "Continue"

$root       = Join-Path $env:LOCALAPPDATA "OutlookAiAssistant"
$serviceDir = Join-Path $root "embedding_service"
$appDir     = Join-Path $root "app"
$logsDir    = Join-Path $root "logs"
$addinDll   = Join-Path $appDir "OutlookAiAssistant.dll"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $OutputPath = Join-Path ([Environment]::GetFolderPath("Desktop")) `
    "OutlookAiAssistant-Diagnose.txt"
}

$lines = New-Object System.Collections.Generic.List[string]

function Add-Line {
  param([string]$Text = "")
  $lines.Add($Text)
  Write-Host $Text
}

function Add-Header {
  param([string]$Title)
  Add-Line ""
  Add-Line ("=" * 70)
  Add-Line $Title
  Add-Line ("=" * 70)
}

function Add-Path-Status {
  param([string]$Label, [string]$Path)
  if (Test-Path -LiteralPath $Path) {
    Add-Line ("[OK]   {0}: {1}" -f $Label, $Path)
  }
  else {
    Add-Line ("[缺失] {0}: {1}" -f $Label, $Path)
  }
}

Add-Line "Outlook AI Assistant 诊断报告"
Add-Line ("生成时间: {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Add-Line ("计算机: {0}   用户: {1}" -f $env:COMPUTERNAME, $env:USERNAME)
Add-Line ("PowerShell: {0}" -f $PSVersionTable.PSVersion)

# ---------------------------------------------------------------------
Add-Header "1. 关键路径是否存在"
Add-Path-Status -Label "数据根目录"     -Path $root
Add-Path-Status -Label "加载项目录"     -Path $appDir
Add-Path-Status -Label "加载项 DLL"     -Path $addinDll
Add-Path-Status -Label "语义服务目录"   -Path $serviceDir
Add-Path-Status -Label "服务启动脚本"   -Path (Join-Path $serviceDir "start-service.cmd")
Add-Path-Status -Label "服务 venv"      -Path (Join-Path $serviceDir ".venv\Scripts\python.exe")
Add-Path-Status -Label "模型目录"       -Path (Join-Path $root "models")
Add-Path-Status -Label "索引目录"       -Path (Join-Path $root "semantic-index")
Add-Path-Status -Label "日志目录"       -Path $logsDir

# ---------------------------------------------------------------------
Add-Header "2. 加载项 DLL 版本（期望 0.6.7 或更高）"
if (Test-Path -LiteralPath $addinDll) {
  try {
    $ver = (Get-Item -LiteralPath $addinDll).VersionInfo
    Add-Line ("FileVersion    : {0}" -f $ver.FileVersion)
    Add-Line ("ProductVersion : {0}" -f $ver.ProductVersion)
    Add-Line ("修改时间       : {0}" -f (Get-Item -LiteralPath $addinDll).LastWriteTime)
  }
  catch {
    Add-Line ("读取版本失败: {0}" -f $_.Exception.Message)
  }
}
else {
  Add-Line "未找到加载项 DLL —— 加载项可能从未安装成功。"
}

# ---------------------------------------------------------------------
Add-Header "3. 语义服务目录内容"
if (Test-Path -LiteralPath $serviceDir) {
  Get-ChildItem -LiteralPath $serviceDir -Force |
    Select-Object Mode, LastWriteTime, Length, Name |
    Format-Table -AutoSize |
    Out-String -Width 200 |
    ForEach-Object { Add-Line $_ }
}
else {
  Add-Line "服务目录不存在。若你曾手动运行过 start-service.cmd，"
  Add-Line "它很可能来自解压目录，而不是这里的安装目录。"
}

# ---------------------------------------------------------------------
Add-Header "4. 服务启动日志 service-start.log（0.6.7 新增）"
$startLog = Join-Path $logsDir "service-start.log"
if (Test-Path -LiteralPath $startLog) {
  Add-Line ("文件: {0}" -f $startLog)
  Add-Line ("大小: {0} 字节" -f (Get-Item -LiteralPath $startLog).Length)
  Add-Line "--- 全部内容 ---"
  Get-Content -LiteralPath $startLog | ForEach-Object { Add-Line $_ }
}
else {
  Add-Line "不存在。"
  Add-Line "=> 说明'启动服务'按钮从未真正执行到新逻辑，"
  Add-Line "   加载项 DLL 很可能仍是旧版本。"
}

# ---------------------------------------------------------------------
Add-Header "5. 语义服务日志 semantic-service.log（最后 40 行）"
$svcLog = Join-Path $logsDir "semantic-service.log"
if (Test-Path -LiteralPath $svcLog) {
  Add-Line ("文件: {0}" -f $svcLog)
  Add-Line ("大小: {0} 字节" -f (Get-Item -LiteralPath $svcLog).Length)
  Get-Content -LiteralPath $svcLog -Tail 40 | ForEach-Object { Add-Line $_ }
}
else {
  Add-Line "不存在。"
}

# ---------------------------------------------------------------------
Add-Header "6. 加载项日志 addin.log（最后 40 行）"
$addinLog = Join-Path $logsDir "addin.log"
if (Test-Path -LiteralPath $addinLog) {
  Add-Line ("文件: {0}" -f $addinLog)
  Add-Line ("大小: {0} 字节" -f (Get-Item -LiteralPath $addinLog).Length)
  Get-Content -LiteralPath $addinLog -Tail 40 | ForEach-Object { Add-Line $_ }
}
else {
  Add-Line "不存在。"
}

# ---------------------------------------------------------------------
Add-Header "7. 正在运行的服务进程"
$procs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
  Where-Object {
    ($_.Name -in @("python.exe", "pythonw.exe") -and
    $_.CommandLine -and
    $_.CommandLine -like "*app.main*") -or
    ($_.Name -eq "cmd.exe" -and
    $_.CommandLine -and
    $_.CommandLine -like "*start-service.cmd*")
  }
if ($procs) {
  foreach ($p in $procs) {
    Add-Line ("PID {0}  启动于 {1}" -f $p.ProcessId, $p.CreationDate)
    Add-Line ("  命令行: {0}" -f $p.CommandLine)
  }
}
else {
  Add-Line "没有正在运行的服务进程。"
}

# ---------------------------------------------------------------------
Add-Header "8. 端口 8765 监听状态"
try {
  $conn = Get-NetTCPConnection -LocalPort 8765 -ErrorAction SilentlyContinue
  if ($conn) {
    $conn | Select-Object LocalAddress, LocalPort, State, OwningProcess |
      Format-Table -AutoSize | Out-String -Width 200 |
      ForEach-Object { Add-Line $_ }
  }
  else {
    Add-Line "端口 8765 当前没有监听。"
  }
}
catch {
  Add-Line ("查询端口失败: {0}" -f $_.Exception.Message)
}

# ---------------------------------------------------------------------
Add-Header "9. /health 实际返回"
try {
  $resp = Invoke-WebRequest -Uri "http://127.0.0.1:8765/health" `
    -UseBasicParsing -TimeoutSec 5
  Add-Line ("HTTP 状态: {0}" -f $resp.StatusCode)
  Add-Line "--- 响应正文 ---"
  Add-Line $resp.Content
}
catch {
  Add-Line ("请求失败: {0}" -f $_.Exception.Message)
}

# ---------------------------------------------------------------------
Add-Header "10. Python 环境"
foreach ($cmd in @("py", "python")) {
  try {
    $out = & $cmd --version 2>&1
    Add-Line ("{0} --version  ->  {1}" -f $cmd, ($out -join " "))
  }
  catch {
    Add-Line ("{0} 不可用" -f $cmd)
  }
}
try {
  $pyList = & py -0 2>&1
  Add-Line "--- py -0 ---"
  $pyList | ForEach-Object { Add-Line $_ }
}
catch {
  Add-Line "py 启动器不可用"
}
$venvPython = Join-Path $serviceDir ".venv\Scripts\python.exe"
if (Test-Path -LiteralPath $venvPython) {
  Add-Line "--- 虚拟环境 PyTorch 状态 ---"
  try {
    $torchInfo = & $venvPython -c "import torch; print('  torch: ' + torch.__version__); print('  CUDA available: ' + str(torch.cuda.is_available())); print('  device count: ' + str(torch.cuda.device_count() if torch.cuda.is_available() else 0)); print('  device name: ' + (torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'N/A'))" 2>&1
    $torchInfo | ForEach-Object { Add-Line $_ }
  }
  catch {
    Add-Line ("  检查 torch 失败: {0}" -f $_.Exception.Message)
  }
}

# ---------------------------------------------------------------------
Add-Header "11. NVIDIA GPU 检测"
$smiPath = $null
$candidates = @(
  "$env:SystemRoot\System32\nvidia-smi.exe",
  "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
  "${env:ProgramW6432}\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
)
foreach ($c in $candidates) {
  if (Test-Path -LiteralPath $c) { $smiPath = $c; break }
}
if (-not $smiPath) {
  $smi = Get-Command nvidia-smi -ErrorAction SilentlyContinue
  if ($smi) { $smiPath = $smi.Source }
}
if ($smiPath) {
  Add-Line ("nvidia-smi: {0}" -f $smiPath)
  try {
    & $smiPath -L 2>&1 | ForEach-Object { Add-Line $_ }
  }
  catch {
    Add-Line ("执行 nvidia-smi 失败: {0}" -f $_.Exception.Message)
  }
}
else {
  Add-Line "未找到 nvidia-smi（无 NVIDIA 驱动，或不在已知路径中）。"
}
try {
  $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
  Add-Line ("显卡: {0}" -f ($gpus -join "; "))
}
catch { }

# ---------------------------------------------------------------------
Add-Header "12. 已安装的 Outlook 加载项注册表项"
try {
  $key = "HKCU:\Software\Microsoft\Office\Outlook\Addins\OutlookAiAssistant.Connect"
  if (Test-Path $key) {
    Get-ItemProperty $key | Format-List | Out-String -Width 200 |
      ForEach-Object { Add-Line $_ }
  }
  else {
    Add-Line "注册表项不存在 —— 加载项未注册。"
  }
}
catch {
  Add-Line ("读取注册表失败: {0}" -f $_.Exception.Message)
}

# ---------------------------------------------------------------------
Add-Header "13. 解压目录中是否有另一份服务（手动启动的来源）"
$candidates = @()
$searchRoots = @(
  (Join-Path $env:USERPROFILE "Downloads"),
  (Join-Path $env:USERPROFILE "Desktop"),
  (Join-Path $env:USERPROFILE "Documents")
)
foreach ($sr in $searchRoots) {
  if (Test-Path -LiteralPath $sr) {
    $found = Get-ChildItem -LiteralPath $sr -Recurse -Depth 3 `
      -Filter "start-service.cmd" -ErrorAction SilentlyContinue
    foreach ($f in $found) { $candidates += $f.FullName }
  }
}
if ($candidates.Count -gt 0) {
  Add-Line "在以下位置发现 start-service.cmd："
  $candidates | ForEach-Object { Add-Line ("  " + $_) }
  Add-Line ""
  Add-Line "如果这里有一份，而你手动成功启动过，说明你启动的是解压目录里的服务，"
  Add-Line "而不是安装目录里的服务。"
}
else {
  Add-Line "未在 Downloads/Desktop/Documents 中发现 start-service.cmd。"
}

# ---------------------------------------------------------------------
Add-Header "14. 结论与建议"

$dllPath = $addinDll
$dllVersion = ""
if (Test-Path -LiteralPath $dllPath) {
  $dllVersion = (Get-Item -LiteralPath $dllPath).VersionInfo.FileVersion
}

if (-not (Test-Path -LiteralPath $serviceDir)) {
  Add-Line "[问题] 语义服务没有安装到 $serviceDir"
  Add-Line "       影响：插件里的'启动服务'/'下载模型'/'建立索引'都无法工作。"
  Add-Line "       解决：解压最新发布包，在 embedding_service 目录运行"
  Add-Line "             install-service.cmd"
  Add-Line ""
}
else {
  Add-Line "[OK]   语义服务已安装。"
}

if ($dllVersion -and ($dllVersion -notmatch '^0\.(6\.[89]|[7-9])')) {
  Add-Line "[问题] 加载项 DLL 版本是 $dllVersion，不是最新版。"
  Add-Line "       影响：新修复不会生效，行为与旧版完全一样。"
  Add-Line "       解决：完全退出 Outlook，对最新发布包运行 Install.cmd。"
  Add-Line ""
}
elseif ($dllVersion) {
  Add-Line "[OK]   加载项 DLL 版本: $dllVersion"
}

if (-not (Test-Path -LiteralPath $startLog)) {
  Add-Line "[问题] 没有 service-start.log，说明'启动服务'按钮从未调用新逻辑。"
  Add-Line "       如果上面显示 DLL 版本较旧，先更新加载项再重试。"
  Add-Line ""
}

if (-not (Test-Path -LiteralPath (Join-Path $serviceDir ".venv\Scripts\python.exe"))) {
  if (Test-Path -LiteralPath $serviceDir) {
    Add-Line "[问题] 服务已安装但缺少 .venv（依赖未安装）。"
    Add-Line "       解决：在服务目录运行 setup-venv.cmd。"
    Add-Line ""
  }
}

Add-Line "以上[问题]项请按对应'解决'处理，然后重新运行本诊断确认。"

# ---------------------------------------------------------------------
Add-Header "诊断结束"
Add-Line ("报告已写入: {0}" -f $OutputPath)

try {
  $lines -join [Environment]::NewLine |
    Set-Content -LiteralPath $OutputPath -Encoding UTF8
  Write-Host ""
  Write-Host ("报告已保存到: {0}" -f $OutputPath) -ForegroundColor Green
  Write-Host "请把该文件的全部内容发回。" -ForegroundColor Green
}
catch {
  Write-Host ("写入报告失败: {0}" -f $_.Exception.Message) -ForegroundColor Red
}

Write-Host ""
if ($Host.Name -eq "ConsoleHost" -and -not $NoPause) {
  Write-Host "按任意键关闭..." -ForegroundColor Cyan
  try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { }
}
