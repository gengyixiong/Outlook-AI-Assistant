using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using OutlookAiAssistant.Diagnostics;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Starts and stops the local semantic search service as a background
    /// process.
    ///
    /// The service must run outside Outlook: it hosts PyTorch and a FastAPI
    /// server, and Outlook COM cannot be blocked or shared with it. This
    /// launcher only manages the process; it never loads Python in-process.
    ///
    /// The service is started with pythonw.exe and no console window, so it
    /// behaves like a resident background service. Its state is reported
    /// through the local HTTP endpoint, not through a console.
    ///
    /// If a console window is needed for troubleshooting, set the environment
    /// variable OAA_SERVICE_CONSOLE=1 before starting Outlook; the launcher
    /// then uses the visible start-service.cmd path instead.
    /// </summary>
    public sealed class SemanticServiceLauncher
    {
        private readonly SemanticSearchClient _client;
        private Process _startedProcess;

        public SemanticServiceLauncher(SemanticSearchClient client)
        {
            if (client == null)
            {
                throw new ArgumentNullException("client");
            }

            _client = client;
        }

        /// <summary>
        /// True when the service process or its background launcher is still alive.
        /// </summary>
        public bool IsServiceProcessAlive()
        {
            if (_startedProcess != null)
            {
                try
                {
                    if (!_startedProcess.HasExited)
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return IsProcessAlive();
        }

        /// <summary>Directory the service is installed to, if present.</summary>
        public static string InstalledServiceDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookAiAssistant",
                    "embedding_service");
            }
        }

        /// <summary>
        /// Default local data directories. These are deterministic, so the
        /// settings UI can show and open them even when the service is not
        /// running and cannot report its own paths through /health.
        /// </summary>
        public static string ProductDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookAiAssistant");
            }
        }

        public static string DefaultModelDirectory
        {
            get { return Path.Combine(ProductDirectory, "models"); }
        }

        public static string DefaultIndexDirectory
        {
            get { return Path.Combine(ProductDirectory, "semantic-index"); }
        }

        public static string DefaultLogDirectory
        {
            get { return Path.Combine(ProductDirectory, "logs"); }
        }

        /// <summary>
        /// True when a service process (python/pythonw running app.main from
        /// our install dir) currently exists. Used to detect a crash quickly
        /// instead of waiting for the full startup timeout.
        /// </summary>
        public static bool IsProcessAlive()
        {
            try
            {
                // 1. If port 8765 is listening, the service is definitely alive.
                try
                {
                    System.Net.NetworkInformation.IPGlobalProperties properties =
                        System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties();
                    System.Net.IPEndPoint[] endPoints = properties.GetActiveTcpListeners();
                    if (endPoints != null)
                    {
                        for (int i = 0; i < endPoints.Length; i++)
                        {
                            if (endPoints[i].Port == 8765)
                            {
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                }

                // 2. Check for python / pythonw processes.
                Process[] pythonProcesses = Process.GetProcessesByName("python");
                Process[] pythonwProcesses = Process.GetProcessesByName("pythonw");
                int pythonCount = (pythonProcesses != null ? pythonProcesses.Length : 0)
                    + (pythonwProcesses != null ? pythonwProcesses.Length : 0);

                if (pythonCount == 0)
                {
                    return false;
                }

                bool inspectedAny = false;
                Process[] all = new Process[pythonCount];
                int offset = 0;
                if (pythonProcesses != null)
                {
                    Array.Copy(pythonProcesses, 0, all, offset, pythonProcesses.Length);
                    offset += pythonProcesses.Length;
                }
                if (pythonwProcesses != null)
                {
                    Array.Copy(pythonwProcesses, 0, all, offset, pythonwProcesses.Length);
                }

                for (int i = 0; i < all.Length; i++)
                {
                    using (Process process = all[i])
                    {
                        string path = null;
                        try
                        {
                            path = process.MainModule == null
                                ? null
                                : process.MainModule.FileName;
                        }
                        catch
                        {
                            // Access denied or bitness mismatch.
                        }

                        if (!string.IsNullOrEmpty(path))
                        {
                            inspectedAny = true;
                            if (path.IndexOf(
                                InstalledServiceDirectory,
                                StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return true;
                            }
                        }
                    }
                }

                // If all python processes threw on MainModule inspection (e.g. 32-bit Outlook
                // vs 64-bit Python), assume alive to avoid false-positive exit during startup.
                if (!inspectedAny && pythonCount > 0)
                {
                    return true;
                }
            }
            catch
            {
                // If we cannot enumerate, assume it is alive rather than
                // reporting a false crash.
                return true;
            }

            return false;
        }

        /// <summary>
        /// True when the service answers on the loopback endpoint.
        /// </summary>
        public bool IsRunning()
        {
            return _client.GetStatus().Reachable;
        }

        /// <summary>True when the service code and its venv are installed.</summary>
        public static bool IsInstalled()
        {
            try
            {
                string directory = InstalledServiceDirectory;
                string python = Path.Combine(
                    directory, ".venv", "Scripts", "python.exe");
                return File.Exists(PythonPath()) && File.Exists(python);
            }
            catch
            {
                return false;
            }
        }

        private static string PythonPath()
        {
            return Path.Combine(
                InstalledServiceDirectory, ".venv", "Scripts", "python.exe");
        }

        /// <summary>
        /// Builds the cmd.exe arguments that run the start script hidden and
        /// capture its output.
        ///
        /// Exposed for tests because cmd quoting is treacherous: a form that
        /// looks right makes cmd treat the quoted script path as a window
        /// title, run nothing, and exit immediately, so the service never
        /// starts. The working form is `call "script" > "log" 2>&1`, which
        /// also survives spaces in either path.
        /// </summary>
        public static string BuildHiddenStartArguments(
            string scriptPath, string logPath)
        {
            if (string.IsNullOrWhiteSpace(scriptPath))
            {
                throw new ArgumentException("scriptPath is required",
                    "scriptPath");
            }

            if (string.IsNullOrWhiteSpace(logPath))
            {
                throw new ArgumentException("logPath is required", "logPath");
            }

            return "/c call \"" + scriptPath + "\" > \""
                + logPath + "\" 2>&1";
        }

        /// <summary>
        /// Starts the service in the background. Returns null on success,
        /// otherwise a message explaining why it could not be started.
        ///
        /// This launches the same start-service.cmd the user runs by hand,
        /// <summary>Process ID of the launched cmd.exe / service process.</summary>
        public int StartedProcessId
        {
            get
            {
                if (_startedProcess != null)
                {
                    try
                    {
                        return _startedProcess.Id;
                    }
                    catch
                    {
                    }
                }
                return 0;
            }
        }

        /// <summary>
        /// Starts the service in the background without creating a console window or taskbar button.
        /// If OAA_SERVICE_CONSOLE=1 is set, opens start-service.cmd in a visible console for debugging.
        /// </summary>
        public string Start()
        {
            if (IsRunning())
            {
                return null;
            }

            if (!IsInstalled())
            {
                return "尚未安装本地语义搜索服务。请先运行发布包中 "
                    + "embedding_service 目录下的 install-service.cmd。";
            }

            try
            {
                if (string.Equals(
                    Environment.GetEnvironmentVariable("OAA_SERVICE_CONSOLE"),
                    "1",
                    StringComparison.OrdinalIgnoreCase))
                {
                    string script = Path.Combine(
                        InstalledServiceDirectory, "start-service.cmd");

                    ProcessStartInfo consoleInfo = new ProcessStartInfo();
                    consoleInfo.FileName = "cmd.exe";
                    consoleInfo.Arguments = "/c \"" + script + "\"";
                    consoleInfo.WorkingDirectory = InstalledServiceDirectory;
                    consoleInfo.UseShellExecute = true;
                    consoleInfo.WindowStyle = ProcessWindowStyle.Normal;

                    _startedProcess = Process.Start(consoleInfo);
                    return null;
                }

                string pythonExe = PythonPath();
                if (!File.Exists(pythonExe))
                {
                    return "未找到 Python 运行环境: " + pythonExe;
                }

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = pythonExe;
                info.Arguments = "-m app.main";
                info.WorkingDirectory = InstalledServiceDirectory;
                info.CreateNoWindow = true;
                info.UseShellExecute = false;
                info.WindowStyle = ProcessWindowStyle.Hidden;
                info.EnvironmentVariables["PYTHONUNBUFFERED"] = "1";

                _startedProcess = Process.Start(info);
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Starting the semantic service failed.", ex);
                return "无法启动本地语义搜索服务：" + ex.Message;
            }
        }

        /// <summary>
        /// Stops the service. Returns null on success, otherwise a message.
        /// </summary>
        public string Stop()
        {
            if (!IsRunning())
            {
                _startedProcess = null;
                return null;
            }

            try
            {
                // Preferred: ask the service to exit, so it can close the
                // database cleanly.
                _client.ShutdownService();
            }
            catch (Exception ex)
            {
                Logger.Error("Graceful service shutdown failed.", ex);
            }

            // Wait briefly for the port to close.
            for (int attempt = 0; attempt < 20; attempt++)
            {
                Thread.Sleep(250);
                if (!IsRunning())
                {
                    _startedProcess = null;
                    return null;
                }
            }

            // Fallback: terminate the process holding the service.
            string error = KillServiceProcesses();
            for (int attempt = 0; attempt < 20; attempt++)
            {
                Thread.Sleep(250);
                if (!IsRunning())
                {
                    _startedProcess = null;
                    return null;
                }
            }

            return string.IsNullOrWhiteSpace(error)
                ? "服务未能在预期时间内停止，请稍后重试。"
                : error;
        }

        /// <summary>Stops and starts the service again.</summary>
        public string Restart()
        {
            string stopError = Stop();
            if (!string.IsNullOrWhiteSpace(stopError))
            {
                return stopError;
            }

            string startError = Start();
            if (!string.IsNullOrWhiteSpace(startError))
            {
                return startError;
            }

            // Give it a moment so the caller's next status poll is accurate.
            for (int attempt = 0; attempt < 60; attempt++)
            {
                Thread.Sleep(1000);
                if (IsRunning())
                {
                    return null;
                }
            }

            return "服务已启动但尚未响应，请稍后在“刷新状态”中确认。";
        }

        /// <summary>
        /// Builds the PowerShell script that targets service processes and port listeners.
        /// Exposed for testing.
        /// </summary>
        public static string BuildKillPowerShellScript()
        {
            return
                "$ErrorActionPreference = 'SilentlyContinue'; "
                + "Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { "
                + "(($_.Name -eq 'python.exe' -or $_.Name -eq 'pythonw.exe') -and "
                + "($_.CommandLine -like '*-m app.main*' -or $_.CommandLine -like '*embedding_service*')) -or "
                + "($_.Name -eq 'cmd.exe' -and $_.CommandLine -like '*start-service.cmd*') "
                + "} | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; "
                + "Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { "
                + "Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }; exit 0";
        }

        /// <summary>
        /// Encodes the kill script as a base64 UTF-16LE string for PowerShell -EncodedCommand.
        /// This eliminates all command-line quotation and parsing pitfalls.
        /// </summary>
        public static string BuildKillEncodedCommand()
        {
            string script = BuildKillPowerShellScript();
            return Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
        }

        /// <summary>
        /// Kills python processes running our service module. Used only when
        /// the graceful shutdown path fails.
        /// </summary>
        private static string KillServiceProcesses()
        {
            try
            {
                // First try direct .NET process termination for same-bitness processes.
                KillMatchingProcessesByName("python");
                KillMatchingProcessesByName("pythonw");

                // Use PowerShell -EncodedCommand to reliably kill across bitness
                // boundaries (e.g. 32-bit Outlook managing 64-bit Python) and cmd.exe wrapper.
                string encodedCommand = BuildKillEncodedCommand();
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "powershell.exe";
                info.Arguments = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand "
                    + encodedCommand;
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.RedirectStandardError = true;
                using (Process process = Process.Start(info))
                {
                    string err = process.StandardError.ReadToEnd();
                    process.WaitForExit(10000);
                    if (process.ExitCode != 0 && !string.IsNullOrWhiteSpace(err))
                    {
                        Logger.Error("PowerShell kill returned exit code "
                            + process.ExitCode + ": " + err, null);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Terminating the service process failed.", ex);
                return "无法结束服务进程：" + ex.Message;
            }
        }

        private static void KillMatchingProcessesByName(string processName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(processName);
                if (processes == null)
                {
                    return;
                }

                for (int i = 0; i < processes.Length; i++)
                {
                    using (Process p = processes[i])
                    {
                        try
                        {
                            string path = null;
                            try
                            {
                                path = p.MainModule == null
                                    ? null
                                    : p.MainModule.FileName;
                            }
                            catch
                            {
                                // Access denied or bitness mismatch; PowerShell handles this.
                            }

                            if (!string.IsNullOrEmpty(path)
                                && path.IndexOf(
                                    InstalledServiceDirectory,
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                p.Kill();
                            }
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }
        }

        /// <summary>Opens a directory in File Explorer. Returns null on success.</summary>
        public static string OpenDirectory(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return "目录路径未知。请先启动语义服务。";
                }

                // Create it when missing so the button always opens something
                // instead of failing on a not-yet-created folder.
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "explorer.exe";
                info.Arguments = "\"" + path + "\"";
                info.UseShellExecute = true;
                Process.Start(info);
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Opening a folder failed.", ex);
                return "无法打开目录：" + ex.Message;
            }
        }
    }
}
