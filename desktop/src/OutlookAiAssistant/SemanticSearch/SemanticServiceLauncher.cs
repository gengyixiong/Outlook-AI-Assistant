using System;
using System.Diagnostics;
using System.IO;
using OutlookAiAssistant.Diagnostics;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Starts the local semantic search service as a separate process.
    ///
    /// The service must run outside Outlook: it hosts PyTorch and a FastAPI
    /// server, and Outlook COM cannot be blocked or shared with it. This
    /// launcher only starts the process; it never loads Python in-process.
    ///
    /// The start script opens a console window on purpose so the user can see
    /// whether the service actually came up and can stop it by closing the
    /// window.
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
                string script = Path.Combine(directory, "start-service.cmd");
                string python = Path.Combine(
                    directory, ".venv", "Scripts", "python.exe");
                return File.Exists(script) && File.Exists(python);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Starts the service. Returns null on success, otherwise a message
        /// explaining why it could not be started.
        /// </summary>
        public string Start()
        {
            if (IsRunning())
            {
                return null;
            }

            string directory = InstalledServiceDirectory;
            string script = Path.Combine(directory, "start-service.cmd");
            if (!File.Exists(script))
            {
                return "尚未安装本地语义搜索服务。请先运行发布包中 "
                    + "embedding_service 目录下的 install-service.cmd。";
            }

            try
            {
                ProcessStartInfo info = new ProcessStartInfo();
                // cmd /k keeps the console open after the script ends so the
                // user can read any error instead of the window vanishing.
                info.FileName = "cmd.exe";
                info.Arguments = "/c \"" + script + "\"";
                info.WorkingDirectory = directory;
                info.UseShellExecute = true;
                info.WindowStyle = ProcessWindowStyle.Normal;
                Process.Start(info);
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("Starting the semantic service failed.", ex);
                return "无法启动本地语义搜索服务：" + ex.Message;
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
