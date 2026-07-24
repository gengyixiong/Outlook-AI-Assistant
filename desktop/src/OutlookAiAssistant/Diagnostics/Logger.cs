using System;
using System.IO;
using System.Text;

namespace OutlookAiAssistant.Diagnostics
{
    /// <summary>
    /// Local operational log. Callers must never pass mail bodies, prompts or keys.
    /// </summary>
    public static class Logger
    {
        private static readonly object SyncRoot = new object();

        public static void Info(string message)
        {
            Write("INFO", message, null);
        }

        public static void Error(string message, Exception exception)
        {
            Write("ERROR", message, exception);
        }

        private static void Write(string level, string message, Exception exception)
        {
            try
            {
                lock (SyncRoot)
                {
                    string directory = Path.Combine(
                        Environment.GetFolderPath(
                            Environment.SpecialFolder.LocalApplicationData),
                        "OutlookAiAssistant",
                        "logs");
                    Directory.CreateDirectory(directory);
                    string path = Path.Combine(directory, "addin.log");

                    if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024)
                    {
                        string previous = Path.Combine(directory, "addin.previous.log");
                        if (File.Exists(previous))
                        {
                            File.Delete(previous);
                        }

                        File.Move(path, previous);
                    }

                    StringBuilder line = new StringBuilder();
                    line.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    line.Append(" [");
                    line.Append(level);
                    line.Append("] ");
                    line.Append(message ?? string.Empty);
                    if (exception != null)
                    {
                        line.Append(" | ");
                        line.Append(exception.GetType().Name);
                        line.Append(": ");
                        line.Append(exception.Message);
                    }

                    line.AppendLine();
                    File.AppendAllText(path, line.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // Logging must never make Outlook or the add-in fail.
            }
        }
    }
}
