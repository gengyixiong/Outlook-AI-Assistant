using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OutlookAiAssistant.Diagnostics;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Manages the system tray icon in the bottom-right corner of the taskbar
    /// and controls visibility (show/hide) of the semantic service console window.
    /// </summary>
    public sealed class SemanticServiceTray : IDisposable
    {
        private NotifyIcon _notifyIcon;
        private readonly SemanticServiceLauncher _launcher;
        private bool _isConsoleHidden;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const int SW_HIDE = 0;
        private const int SW_SHOWNORMAL = 1;
        private const int SW_RESTORE = 9;

        public event EventHandler ConsoleVisibilityChanged;

        public SemanticServiceTray(SemanticServiceLauncher launcher)
        {
            _launcher = launcher;
        }

        public bool IsConsoleHidden
        {
            get { return _isConsoleHidden; }
        }

        /// <summary>
        /// Finds the HWND of the console window hosting start-service.cmd.
        /// Matches by process ID or window title.
        /// </summary>
        public IntPtr FindConsoleWindowHandle()
        {
            int targetPid = _launcher != null ? _launcher.StartedProcessId : 0;
            IntPtr found = IntPtr.Zero;

            try
            {
                EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
                {
                    if (targetPid > 0)
                    {
                        uint pid;
                        GetWindowThreadProcessId(hWnd, out pid);
                        if (pid == (uint)targetPid)
                        {
                            found = hWnd;
                            return false;
                        }
                    }

                    StringBuilder sb = new StringBuilder(512);
                    GetWindowText(hWnd, sb, 512);
                    string title = sb.ToString();
                    if (!string.IsNullOrEmpty(title) &&
                        (title.IndexOf("Outlook AI Assistant", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         title.IndexOf("start-service.cmd", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         title.IndexOf("语义搜索服务", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        found = hWnd;
                        return false;
                    }

                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                Logger.Error("Error finding console window handle.", ex);
            }

            return found;
        }

        /// <summary>
        /// Hides the console window completely from desktop and taskbar.
        /// </summary>
        public bool HideConsole()
        {
            IntPtr hWnd = FindConsoleWindowHandle();
            if (hWnd != IntPtr.Zero)
            {
                ShowWindow(hWnd, SW_HIDE);
                _isConsoleHidden = true;
                UpdateTrayIcon(true);
                try
                {
                    if (_notifyIcon != null)
                    {
                        _notifyIcon.ShowBalloonTip(
                            2000,
                            "AI 邮件助手",
                            "语义搜索服务控制台已隐藏至右下角系统托盘。\r\n双击托盘图标或在面板中点击“显示控制台”可随时打开。",
                            ToolTipIcon.Info);
                    }
                }
                catch
                {
                }

                if (ConsoleVisibilityChanged != null)
                {
                    ConsoleVisibilityChanged(this, EventArgs.Empty);
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Restores and shows the console window on screen.
        /// </summary>
        public bool ShowConsole()
        {
            IntPtr hWnd = FindConsoleWindowHandle();
            if (hWnd != IntPtr.Zero)
            {
                ShowWindow(hWnd, SW_RESTORE);
                _isConsoleHidden = false;
                UpdateTrayIcon(true);
                if (ConsoleVisibilityChanged != null)
                {
                    ConsoleVisibilityChanged(this, EventArgs.Empty);
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Checks whether the console window is currently visible on screen.
        /// </summary>
        public bool IsConsoleVisible
        {
            get
            {
                IntPtr hWnd = FindConsoleWindowHandle();
                if (hWnd == IntPtr.Zero)
                {
                    return false;
                }
                return IsWindowVisible(hWnd);
            }
        }

        /// <summary>
        /// Toggles between hiding and showing the console window.
        /// </summary>
        public void ToggleConsole()
        {
            if (IsConsoleVisible)
            {
                HideConsole();
            }
            else
            {
                ShowConsole();
            }
        }

        /// <summary>
        /// Updates the state and visibility of the notification area tray icon.
        /// </summary>
        public void UpdateTrayIcon(bool serviceRunning)
        {
            if (!serviceRunning)
            {
                if (_notifyIcon != null)
                {
                    _notifyIcon.Visible = false;
                }
                _isConsoleHidden = false;
                return;
            }

            if (_notifyIcon == null)
            {
                _notifyIcon = new NotifyIcon();
                _notifyIcon.Icon = SystemIcons.Application;
                _notifyIcon.DoubleClick += delegate { ToggleConsole(); };

                ContextMenu menu = new ContextMenu();
                menu.MenuItems.Add(new MenuItem("显示/隐藏控制台窗口", delegate { ToggleConsole(); }));
                menu.MenuItems.Add(new MenuItem("停止服务", delegate
                {
                    if (_launcher != null)
                    {
                        _launcher.Stop();
                        UpdateTrayIcon(false);
                    }
                }));
                _notifyIcon.ContextMenu = menu;
            }

            _notifyIcon.Text = "Outlook AI 语义服务 (运行中)";
            _notifyIcon.Visible = true;
        }

        public void Dispose()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }
        }
    }
}
