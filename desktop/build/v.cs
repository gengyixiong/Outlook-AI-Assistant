using System;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;

public class V
{
    public static void Main()
    {
        string root = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookAiAssistant");
        string dir = Path.Combine(root, "embedding_service");
        string script = Path.Combine(dir, "start-service.cmd");
        string log = Path.Combine(root, "logs", "service-start.log");
        Directory.CreateDirectory(Path.GetDirectoryName(log));
        if (File.Exists(log)) File.Delete(log);

        ProcessStartInfo info = new ProcessStartInfo();
        info.FileName = "cmd.exe";
        info.Arguments = "/c call \"" + script + "\" > \"" + log + "\" 2>&1";
        info.UseShellExecute = false;
        info.CreateNoWindow = true;
        info.WorkingDirectory = dir;
        Console.WriteLine("Arguments = " + info.Arguments);
        Process.Start(info);

        for (int i = 1; i <= 60; i++)
        {
            System.Threading.Thread.Sleep(1000);
            try
            {
                using (TcpClient c = new TcpClient())
                {
                    c.Connect("127.0.0.1", 8765);
                    Console.WriteLine("SERVICE READY after " + i + "s");
                    Console.WriteLine("--- log ---");
                    Console.WriteLine(File.ReadAllText(log));
                    return;
                }
            }
            catch { }
        }
        Console.WriteLine("TIMED OUT");
        Console.WriteLine("--- log ---");
        Console.WriteLine(File.Exists(log) ? File.ReadAllText(log) : "(no log)");
    }
}
