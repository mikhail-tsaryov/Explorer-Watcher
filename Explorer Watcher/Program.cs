using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Timers;
using System.Windows.Forms;
using SHDocVw;
using System.Drawing;
using System.Runtime.InteropServices;

namespace ExplorerTrayWatcher
{

    static class Program
    {
        static NotifyIcon trayIcon;
        static List<string> lastPaths = new List<string>();
        static string saveFilePath = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
        static System.Timers.Timer checkTimer;
        static ToolStripMenuItem restoreItem;
        static int countdown = 20;
        static System.Windows.Forms.Timer uiTimer;
        static bool explorerRunning = false;
        static ContextMenuStrip contextMenu;

        [STAThread]
        static void Main()
        {
            // Close the program if an instance is already running
            if (System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            trayIcon = new NotifyIcon
            {
                Icon = Explorer_Watcher.Properties.Resources.EW,
                Visible = true,
                Text = "Explorer Watcher"
            };

            contextMenu = new ContextMenuStrip();
            BuildContextMenu();

            trayIcon.ContextMenuStrip = contextMenu;

            checkTimer = new System.Timers.Timer(20_000); // 20 seconds
            checkTimer.Elapsed += CheckExplorerWindows;
            checkTimer.Start();

            uiTimer = new System.Windows.Forms.Timer();
            uiTimer.Interval = 1000;
            uiTimer.Tick += (s, e) => UpdateCountdown();
            uiTimer.Start();

            countdown = 20; // Ensure countdown is displayed initially
            CheckExplorerWindows(null, null);

            Application.Run();
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
                    ForeColor = Color.Black
                };
                contextMenu.Items.Add(item);
            }

            contextMenu.Items.Add(new ToolStripSeparator());

            restoreItem = new ToolStripMenuItem($"Restore Windows ({countdown}s)");
            restoreItem.Click += (s, e) => RestoreWindows();
            contextMenu.Items.Add(restoreItem);

            contextMenu.Items.Add("Open Saved List", null, (s, e) => OpenPathsFile());

            var clearItem = new ToolStripMenuItem("Clear Saved List");
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

        // P/Invoke для активации окна
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
                var shellWindows = new SHDocVw.ShellWindows();
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
                            if (IsIconic(hwnd)) // свернуто?
                            {
                                ShowWindow(hwnd, SW_RESTORE); // восстановить
                            }
                            SetForegroundWindow(hwnd); // вывести на передний план
                            return;
                        }
                    }
                }

                // Если окно не найдено, открыть новое
                if (Directory.Exists(path))
                {
                    Process.Start("explorer.exe", path);
                }
            }
            catch
            {
                // Игнорируем ошибки
            }
        }


        static void UpdateCountdown()
        {
            if (!explorerRunning)
            {
                restoreItem.Text = "Restore Windows";
                return;
            }

            countdown--;
            if (countdown <= 0)
            {
                countdown = 20;
            }
            restoreItem.Text = $"Restore Windows ({countdown}s)";
        }
        static void CheckExplorerWindows(object sender, ElapsedEventArgs e)
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
                        catch
                        {
                            path = "";
                        }

                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                        {
                            explorerPaths.Add(path);
                        }
                        else
                        {
                            // Если путь невалиден, но окно explorer есть — считаем Explorer запущенным
                            // и не добавляем путь в список, чтобы не сломать логику.
                        }
                    }
                }

                explorerRunning = foundExplorer;

                if (explorerRunning && explorerPaths.Count > 0)
                {
                    lastPaths = explorerPaths.Distinct().ToList();
                    SaveLastPaths();
                    countdown = 20;
                    trayIcon.ContextMenuStrip.Invoke((MethodInvoker)BuildContextMenu);
                }
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
                        string path = new Uri(window.LocationURL).LocalPath;
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            openPaths.Add(path);
                        }
                    }
                }

                foreach (var path in lastPaths)
                {
                    try
                    {
                        if (Directory.Exists(path) && !openPaths.Contains(path))
                            Process.Start("explorer.exe", path);
                    }
                    catch { }
                }
            }
            catch { }
        }

        static void OpenPathsFile()
        {
            try
            {
                if (File.Exists(saveFilePath))
                {
                    Process.Start("notepad.exe", saveFilePath);
                }
            }
            catch { }
        }

        static void Exit()
        {
            trayIcon.Visible = false;
            checkTimer?.Stop();
            uiTimer?.Stop();
            Application.Exit();
        }

        static void SaveLastPaths()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(saveFilePath));
                File.WriteAllLines(saveFilePath, lastPaths);
            }
            catch { }
        }

        static void LoadLastPaths()
        {
            try
            {
                if (File.Exists(saveFilePath))
                {
                    lastPaths = File.ReadAllLines(saveFilePath).Where(Directory.Exists).ToList();
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
    }
}
