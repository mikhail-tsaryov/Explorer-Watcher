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

namespace Explorer_Watcher
{
    static class Program
    {
        private static readonly Icon iconActive = Explorer_Watcher.Properties.Resources.EW_active;
        private static readonly Icon iconInactive = Explorer_Watcher.Properties.Resources.EW_no_active;
        public static ApplicationSettings AppSettings = new ApplicationSettings();

        private static uint AutoSaveIntervalSeconds = 60;
        private static uint AutoSaveIntervalMilliseconds = AutoSaveIntervalSeconds * 1000;

        static private ToolStripMenuItem intervalMenu = new ToolStripMenuItem("Interval");
        static private ToolStripMenuItem customIntervalItem;

        static NotifyIcon trayIcon;
        static List<string> lastPaths = new List<string>();
        static List<string> lastSavedPaths = new List<string>();
        private static string SaveLocation = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
        static System.Timers.Timer checkTimer;
        static ToolStripMenuItem restoreItem;
        static uint countdown = AutoSaveIntervalSeconds;
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

                UpdateIntervalMenu();
            };

            uiTimer.Start();

            AutoSaveIntervalSeconds = AppSettings.GetParameter<uint>("AutoSaveIntervalSeconds");

            var saveLocation = AppSettings.GetParameter<string>("SaveLocation");

            if (saveLocation == "temp")
            {
                SaveLocation = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
            }
            else if (saveLocation == "program")
            {
                SaveLocation = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last_paths.txt");
            }

            countdown = AutoSaveIntervalSeconds;
            UpdateIntervalMenu();

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
                        catch (UriFormatException)
                        {
                            Debug.WriteLine("Path is empty\r\n");
                        }

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

            // Добавляем меню выбора интервала
            var intervalMenu = new ToolStripMenuItem("Autosave interval");

            void AddPreset(string label, uint seconds)
            {
                var item = new ToolStripMenuItem(label) { Tag = seconds };
                item.CheckOnClick = true;  
                item.Click += (s, e) =>
                {
                    AutoSaveIntervalSeconds = seconds;
                    countdown = AutoSaveIntervalSeconds;
                    UpdateIntervalMenu(); // обновит галочки и надписи
                };
                intervalMenu.DropDownOpening += (sender, e) => UpdateIntervalMenu();
                intervalMenu.DropDownItems.Add(item);
            }

            // Предустановленные интервалы
            AddPreset("30 sec", 30);
            AddPreset("1 min", 60);
            AddPreset("2 min", 120);
            AddPreset("5 min", 300);

            intervalMenu.DropDownItems.Add(new ToolStripSeparator());
            /*
            // Кастомный интервал
            customIntervalItem = new ToolStripMenuItem("Custom: ...");
            customIntervalItem.Click += (s, e) =>
            {
                IntervalForm intervalForm = new IntervalForm(AutoSaveIntervalSeconds);
                using var form = intervalForm;
                if (form.ShowDialog() == DialogResult.OK)
                {
                    AutoSaveIntervalSeconds = form.IntervalSeconds;
                    countdown = AutoSaveIntervalSeconds;
                    UpdateIntervalMenu();
                }
            };
            intervalMenu.DropDownItems.Add(customIntervalItem);*/

            contextMenu.Items.Add(intervalMenu);

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
            UpdateIntervalMenu();
        }

        static void UpdateIntervalMenu()
        {
            // Обновляем предустановки
            foreach (ToolStripMenuItem item in intervalMenu.DropDownItems)
            {
                if (item.Tag is uint seconds)
                {
                    item.Checked = (seconds == AutoSaveIntervalSeconds);
                }
            }
            /*
            // Обновляем кастомный пункт (если используется)
            if (customIntervalItem != null)
            {
                customIntervalItem.Text = $"Custom: {FormatSeconds(AutoSaveIntervalSeconds)}";
                customIntervalItem.Checked = !intervalMenu.DropDownItems
                    .OfType<ToolStripMenuItem>()
                    .Any(i => i.Checked);
            }*/
        }
        private static string FormatSeconds(uint seconds)
        {
            return seconds < 60
                ? $"{seconds} sec"
                : $"{seconds / 60} min";
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
                        catch (UriFormatException) { Debug.WriteLine("Path is empty\r\n"); }

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
            catch (UriFormatException) { Debug.WriteLine("Path is empty\r\n"); }
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
            trayIcon.Text = $"ExplorerWatcher: Autosave will be in {countdown}s. Double click to save";

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
                        catch (UriFormatException)
                        {
                            Debug.WriteLine("Path is empty\r\n");
                        }

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
            catch (UriFormatException) 
            { 
                Debug.WriteLine("Path is empty\r\n");
            }
        }

        static void SaveLastPaths()
        {
            try
            {
                var dir = Path.GetDirectoryName(SaveLocation);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllLines(SaveLocation, lastPaths);
            }
            catch (UriFormatException) { Debug.WriteLine("Path is empty\r\n"); }
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
                        catch (UriFormatException)
                        {
                            Debug.WriteLine("Path is empty\r\n");
                        }

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
                if (File.Exists(SaveLocation))
                {
                    lastPaths = File.ReadAllLines(SaveLocation).Where(line => !string.IsNullOrWhiteSpace(line)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
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
                if (!File.Exists(SaveLocation))
                    File.WriteAllLines(SaveLocation, new string[0]);

                Process.Start("notepad.exe", SaveLocation);
            }
            catch (UriFormatException) { Debug.WriteLine("Path is empty\r\n"); }
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
