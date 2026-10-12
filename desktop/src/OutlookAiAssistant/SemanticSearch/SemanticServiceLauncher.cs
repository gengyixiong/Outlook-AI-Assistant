using System;
using System.Diagnostics;
using System.IO;
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

        public SemanticServiceLauncher(SemanticSearchClient client)
        {
            if (client == null)
            {
                throw new ArgumentNullException("client");
            }

            _client = client;
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

        private static string PythonWindowlessPath()
        {
            string path = Path.Combine(
                InstalledServiceDirectory, ".venv", "Scripts", "pythonw.exe");
            return File.Exists(path) ? path : null;
        }

        /// <summary>
        /// Starts the service in the background. Returns null on success,
        /// otherwise a message explaining why it could not be started.
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
                bool wantConsole = string.Equals(
                    Environment.GetEnvironmentVariable("OAA_SERVICE_CONSOLE"),
                    "1",
                    StringComparison.Ordinal);

                ProcessStartInfo info = new ProcessStartInfo();
                if (wantConsole)
                {
                    // Troubleshooting path: visible console running the script.
                    info.FileName = "cmd.exe";
                    info.Arguments = "/c \""
                        + Path.Combine(InstalledServiceDirectory,
                            "start-service.cmd") + "\"";
                    info.UseShellExecute = true;
                    info.WindowStyle = ProcessWindowStyle.Normal;
                }
                else
                {
                    // Resident path: no window at all.
                    string windowless = PythonWindowlessPath();
                    if (windowless != null)
                    {
                        info.FileName = windowless;
                    }
                    else
                    {
                        // Fall back to python.exe but keep the window hidden.
                        info.FileName = PythonPath();
                        info.CreateNoWindow = true;
                    }

                    info.Arguments = "-m app.main";
                    info.UseShellExecute = false;
                    info.CreateNoWindow = true;
                    info.WorkingDirectory = InstalledServiceDirectory;
                    // Make sure the child gets its own environment so it can
                    // set HF_HOME etc. without inheriting stale values.
                    info.EnvironmentVariables["PYTHONUNBUFFERED"] = "1";
                }

                Process.Start(info);
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
                return null;
            }

            try
            {
                // Preferred: ask the service to exit, so it can close the
                // database cleanly.
                _client.ShutdownService();
            }
            catch (SemanticServiceException ex)
            {
                Logger.Error("Graceful service shutdown failed.", ex);
            }

            // Wait briefly for the port to close.
            for (int attempt = 0; attempt < 20; attempt++)
            {
                Thread.Sleep(250);
                if (!IsRunning())
                {
                    return null;
                }
            }

            // Fallback: terminate the process holding the service.
            string error = KillServiceProcesses();
            for (int attempt = 0; attempt < 8; attempt++)
            {
                Thread.Sleep(250);
                if (!IsRunning())
                {
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
        /// Kills python processes running our service module. Used only when
        /// the graceful shutdown path fails.
        /// </summary>
        private static string KillServiceProcesses()
        {
            try
            {
                // Target only python processes started from our install dir
                // running app.main, so unrelated Python work is not touched.
                string query =
                    "Get-CimInstance Win32_Process -Filter \"Name='python.exe' "
                    + "OR Name='pythonw.exe'\" | "
                    + "Where-Object { $_.CommandLine -like '*app.main*' -and "
                    + "$_.CommandLine -like '*embedding_service*' } | "
                    + "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }";

                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "powershell.exe";
                info.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \""
                    + query + "\"";
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                using (Process process = Process.Start(info))
                {
                    process.WaitForExit(10000);
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Terminating the service process failed.", ex);
                return "无法结束服务进程：" + ex.Message;
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
