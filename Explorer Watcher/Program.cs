using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using SHDocVw;
using System.Drawing;

namespace ExplorerWatcher
{
    static class Program
    {
        private static readonly Icon iconActive = Explorer_Watcher.Properties.Resources.EW_active;
        private static readonly Icon iconInactive = Explorer_Watcher.Properties.Resources.EW_no_active;

        private const int AutoSaveIntervalSeconds = 60;
        private const int AutoSaveIntervalMilliseconds = AutoSaveIntervalSeconds * 1000;

        static NotifyIcon trayIcon;
        static List<string> lastPaths = new List<string>();
        static List<string> lastSavedPaths = new List<string>();
        private static readonly string saveFilePath = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
        static System.Timers.Timer checkTimer;
        static ToolStripMenuItem restoreItem;
        static int countdown = AutoSaveIntervalSeconds;
        static readonly bool countdownActive = false;
        static System.Windows.Forms.Timer uiTimer;
        static bool explorerRunning = false;
        static ContextMenuStrip contextMenu;
        static ToolStripMenuItem countdownMenuItem;

        [STAThread]
        static void Main()
        {
            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
                return;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            trayIcon = new NotifyIcon
            {
                Icon = iconInactive,
                Visible = true,
                Text = "Explorer Watcher"
            };

            trayIcon.MouseDoubleClick += NotifyIcon_MouseDoubleClick;

            contextMenu = new ContextMenuStrip();
            trayIcon.ContextMenuStrip = contextMenu;

            trayIcon.ContextMenuStrip.Opening += (s, e) =>
            {
                BuildContextMenu();
            };

            checkTimer = new System.Timers.Timer(AutoSaveIntervalMilliseconds);
            checkTimer.Elapsed += CheckExplorerWindows;
            checkTimer.Start();

            uiTimer = new System.Windows.Forms.Timer
            {
                Interval = 1000
            };
            uiTimer.Tick += (s, e) =>
            {
                if (!explorerRunning)
                {
                    countdown = 0;
                    UpdateIcon();
                    UpdateCountdown();
                    return;
                }

                countdown--;

                if (countdown <= 0)
                {
                    UpdateExplorerWindows();
                }
                else
                {
                    UpdateCountdown();
                    UpdateTooltip();
                }

                UpdateIcon();
            };

            uiTimer.Start();

            countdown = AutoSaveIntervalSeconds;

            InitExplorerWindowWatcher();

            InitialCheck();

            Application.Run();
        }

        static void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (explorerRunning)
            {
                SaveCurrentExplorerWindows();
                trayIcon.BalloonTipTitle = "Save current Explorer windows";
                trayIcon.BalloonTipText = "Saving completed successfully!";
                trayIcon.BalloonTipIcon = ToolTipIcon.Info;
                trayIcon.ShowBalloonTip(3000);
            }
            else
            {
                RestoreWindows();
            }
        }

        static void SaveCurrentExplorerWindows()
        {
            try
            {
                var shellWindows = new SHDocVw.ShellWindows();
                var explorerPaths = new List<string>();

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string path = "";
                        try
                        {
                            path = new Uri(window.LocationURL).LocalPath;
                        }
                        catch { }

                        if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                        {
                            explorerPaths.Add(path);
                        }
                    }
                }

                lastPaths = explorerPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                SaveLastPaths();

                BuildContextMenu();
                UpdateTooltip();
            }
            catch
            {

            }
        }

        static void InitialCheck()
        {
            try
            {
                var shellWindows = new ShellWindows();
                var explorerPaths = new List<string>();
                bool foundExplorer = false;

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        foundExplorer = true;
                        string url = window.LocationURL;
                        string path = "";

                        try
                        {
                            path = new Uri(url).LocalPath;
                        }
                        catch { path = ""; }

                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                            explorerPaths.Add(path);
                    }
                }

                explorerRunning = foundExplorer;
                if (explorerRunning && explorerPaths.Count > 0)
                {
                    lastPaths = explorerPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    lastSavedPaths = new List<string>(lastPaths);
                    SaveLastPaths();
                }
            }
            catch
            {
                explorerRunning = false;
            }

            BuildContextMenu();
            UpdateCountdown();
        }

        static void BuildContextMenu()
        {
            contextMenu.Items.Clear();

            LoadLastPaths();

            if (lastPaths.Count > 0)
            {
                foreach (var path in lastPaths)
                {
                    string folderName = Path.GetFileName(path.TrimEnd(Path.DirectorySeparatorChar));
                    if (string.IsNullOrEmpty(folderName))
                    {
                        try
                        {
                            DriveInfo drive = new DriveInfo(path);
                            folderName = $"Drive {drive.Name.TrimEnd(Path.DirectorySeparatorChar)}";
                        }
                        catch
                        {
                            folderName = path;
                        }
                    }

                    var item = new ToolStripMenuItem(folderName)
                    {
                        ToolTipText = path,
                        Enabled = true,
                        ForeColor = Color.Black
                    };
                    item.Click += (s, e) => OpenOrActivateExplorerWindow(path);
                    contextMenu.Items.Add(item);
                }
            }
            else
            {
                var item = new ToolStripMenuItem("(no saved paths)")
                {
                    Enabled = false,
                    ForeColor = Color.Gray
                };
                contextMenu.Items.Add(item);
            }

            contextMenu.Items.Add(new ToolStripSeparator());

            countdownMenuItem = new ToolStripMenuItem($"Autosave will be in {countdown}s")
            {
                Enabled = false,
                ForeColor = Color.Gray
            };
            contextMenu.Items.Add(countdownMenuItem);

            contextMenu.Items.Add(new ToolStripSeparator());

            restoreItem = new ToolStripMenuItem("Restore windows");
            restoreItem.Click += (s, e) => RestoreWindows();
            contextMenu.Items.Add(restoreItem);

            contextMenu.Items.Add("Open saved list", null, (s, e) => OpenPathsFile());

            var clearItem = new ToolStripMenuItem("Clear saved list");
            clearItem.Click += (s, e) =>
            {
                lastPaths.Clear();
                SaveLastPaths();
                BuildContextMenu();
            };
            contextMenu.Items.Add(clearItem);

            contextMenu.Items.Add("Exit", null, (s, e) => Exit());

            UpdateCountdown();
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_RESTORE = 9;

        static void OpenOrActivateExplorerWindow(string path)
        {
            try
            {
                var shellWindows = new ShellWindows();
                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string windowPath = "";
                        try
                        {
                            windowPath = new Uri(window.LocationURL).LocalPath;
                        }
                        catch { }

                        if (string.Equals(windowPath.TrimEnd(Path.DirectorySeparatorChar), path.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
                        {
                            IntPtr hwnd = new IntPtr(window.HWND);
                            if (IsIconic(hwnd))
                                ShowWindow(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
                            return;
                        }
                    }
                }

                if (Directory.Exists(path))
                    Process.Start("explorer.exe", path);
            }
            catch { }
        }

        static void UpdateCountdown()
        {
            if (countdownMenuItem == null)
                return;

            if (!explorerRunning)
            {
                countdownMenuItem.Text = "No Explorer windows";
                trayIcon.Text = "Explorer Watcher: Double click to restore Explorer windows";
                UpdateIcon();
                return;
            }

            countdownMenuItem.Text = $"Next autosave will be in {countdown}s";
            trayIcon.Text = $"Explorer Watcher: Autosave will be in {countdown}s. Double click to save";

            UpdateIcon();
        }

        static void CheckExplorerWindows(object sender, ElapsedEventArgs e)
        {
            try
            {
                var shellWindows = new ShellWindows();
                var currentPaths = new List<string>();
                bool explorerFound = false;

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        explorerFound = true;
                        string url = window.LocationURL;
                        string path = "";
                        try
                        {
                            path = new Uri(url).LocalPath;
                        }
                        catch { path = ""; }

                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                            currentPaths.Add(path);
                    }
                }

                countdown--;

                if (countdown <= 0)
                {
                    if (explorerFound)
                    {
                        var distinctCurrent = currentPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                        if (!distinctCurrent.SequenceEqual(lastPaths, StringComparer.OrdinalIgnoreCase))
                        {
                            lastPaths = distinctCurrent;
                            SaveLastPaths();
                        }
                    }

                    countdown = AutoSaveIntervalSeconds;
                }

                explorerRunning = explorerFound;
                UpdateTooltip();
            }
            catch
            {
                explorerRunning = false;
            }
        }

        static void RestoreWindows()
        {
            try
            {
                var shellWindows = new ShellWindows();
                var openPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string path = "";
                        try
                        {
                            path = new Uri(window.LocationURL).LocalPath;
                        }
                        catch { }

                        if (!string.IsNullOrWhiteSpace(path))
                            openPaths.Add(path.TrimEnd(Path.DirectorySeparatorChar));
                    }
                }

                foreach (var path in lastPaths)
                {
                    if (!openPaths.Contains(path.TrimEnd(Path.DirectorySeparatorChar)) && Directory.Exists(path))
                    {
                        Process.Start("explorer.exe", path);
                    }
                }
            }
            catch { }
        }

        static void SaveLastPaths()
        {
            try
            {
                var dir = Path.GetDirectoryName(saveFilePath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllLines(saveFilePath, lastPaths);
            }
            catch { }
        }

        static void UpdateExplorerWindows()
        {
            if (!explorerRunning)
            {
                countdown = 0;
                return;
            }

            try
            {
                var shellWindows = new ShellWindows();
                var currentPaths = new List<string>();

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string path = "";
                        try
                        {
                            path = new Uri(window.LocationURL).LocalPath;
                        }
                        catch { }

                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                            currentPaths.Add(path);
                    }
                }

                currentPaths = currentPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                bool pathsChanged = !currentPaths.SequenceEqual(lastSavedPaths, StringComparer.OrdinalIgnoreCase);

                if (pathsChanged)
                {
                    lastSavedPaths = currentPaths;
                    lastPaths = new List<string>(lastSavedPaths);
                    SaveLastPaths();

                    trayIcon?.ContextMenuStrip?.Invoke((MethodInvoker)(() =>
                    {
                        BuildContextMenu();
                        UpdateTooltip();
                    }));
                }

                countdown = AutoSaveIntervalSeconds;
            }
            catch
            {
                explorerRunning = false;
                countdown = 0;
                trayIcon?.ContextMenuStrip?.Invoke((MethodInvoker)(() =>
                {
                    BuildContextMenu();
                    UpdateTooltip();
                }));
            }
        }

        static void LoadLastPaths()
        {
            try
            {
                if (File.Exists(saveFilePath))
                {
                    lastPaths = File.ReadAllLines(saveFilePath).Where(line => !string.IsNullOrWhiteSpace(line)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                }
                else
                {
                    lastPaths = new List<string>();
                }
            }
            catch
            {
                lastPaths = new List<string>();
            }
        }

        static void UpdateTooltip()
        {
            if (!explorerRunning)
                trayIcon.Icon = iconInactive;
            else
                trayIcon.Icon = iconActive;
        }

        static void UpdateIcon()
        {
            trayIcon.Icon = explorerRunning ? iconActive : iconInactive;
        }

        static void OpenPathsFile()
        {
            try
            {
                if (!File.Exists(saveFilePath))
                    File.WriteAllLines(saveFilePath, new string[0]);

                Process.Start("notepad.exe", saveFilePath);
            }
            catch { }
        }

        static void Exit()
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        #region WinEventHook to watch Explorer windows open/close

        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        static IntPtr winEventHook;
        static WinEventDelegate winEventProcDelegate;

        const uint EVENT_OBJECT_CREATE = 0x8000;
        const uint EVENT_OBJECT_DESTROY = 0x8001;
        const uint WINEVENT_OUTOFCONTEXT = 0;

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
            hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
            uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll")]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hWnd);

        static void InitExplorerWindowWatcher()
        {
            winEventProcDelegate = new WinEventDelegate(WinEventProc);
            winEventHook = SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_DESTROY,
                IntPtr.Zero, winEventProcDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
        }

        static void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
            int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hwnd == IntPtr.Zero) return;
            if (idObject != 0 || idChild != 0) return;
            if (!IsWindow(hwnd)) return;

            StringBuilder className = new StringBuilder(256);
            GetClassName(hwnd, className, 256);
            string cls = className.ToString();

            if (cls == "CabinetWClass" || cls == "ExploreWClass")
            {
                bool anyExplorerWindow = false;
                try
                {
                    var shellWindows = new ShellWindows();
                    foreach (InternetExplorer window in shellWindows)
                    {
                        if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                        {
                            anyExplorerWindow = true;
                            break;
                        }
                    }
                }
                catch
                {
                    anyExplorerWindow = false;
                }

                explorerRunning = anyExplorerWindow;
            }
        }

        #endregion
    }
}
