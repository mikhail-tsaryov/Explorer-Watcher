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

        private static int AutoSaveIntervalSeconds = 60;
        private static int AutoSaveIntervalMilliseconds = AutoSaveIntervalSeconds * 1000;

        static ToolStripMenuItem intervalMenu = new ToolStripMenuItem("Autosave interval");
        static ToolStripMenuItem saveLocationMenu = new ToolStripMenuItem("Save directory");

        static NotifyIcon trayIcon;
        static List<string> lastPaths = new List<string>();
        static List<string> lastSavedPaths = new List<string>();
        private static readonly int MaxHistoryFiles = 10;
        private static string SaveDirectory => Path.GetDirectoryName(SaveLocation);
        private static string SaveLocation = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "paths.txt");
        static System.Timers.Timer checkTimer;
        static ToolStripMenuItem restoreItem;
        static int countdown = AutoSaveIntervalSeconds;
        static System.Windows.Forms.Timer uiTimer;
        static bool explorerRunning = false;
        static ContextMenuStrip contextMenu;
        static ToolStripMenuItem countdownMenuItem;
        static Icon currentDynamicIcon = null;
        static bool showProgressBar = true; // по умолчанию включено


        static ToolTip historyToolTip = new ToolTip
        {
            AutoPopDelay = 20000,
            InitialDelay = 400,
            ReshowDelay = 200,
            ShowAlways = true
        };


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
                }

                UpdateIcon();
            };

            uiTimer.Start();

            AutoSaveIntervalSeconds = AppSettings.GetParameter<int>("AutoSaveIntervalSeconds");

            var saveLocation = AppSettings.GetParameter<string>("SaveLocation");

            if (saveLocation == "temp")
            {
                SaveLocation = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
            }
            else if (saveLocation == "program")
            {
                SaveLocation = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last_paths.txt");
            }

            intervalMenu.DropDownOpening += (s, e) => UpdateIntervalMenu();
            saveLocationMenu.DropDownOpening += (s, e) => UpdateSaveLocationMenu();
            countdown = AutoSaveIntervalSeconds;
            UpdateIntervalMenu();
            UpdateSaveLocationMenu();

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

        static string BuildPreviewText(string filePath)
        {
            try
            {
                var lines = File.ReadAllLines(filePath)
                                .Where(l => !string.IsNullOrWhiteSpace(l))
                                .ToList();

                if (lines.Count == 0)
                    return "(empty)";

                var sb = new StringBuilder();

                for (int i = 0; i < lines.Count; i++)
                {
                    sb.AppendLine(lines[i]);
                }

                return sb.ToString();
            }
            catch
            {
                return "(failed to read file)";
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
                UpdateIntervalMenu();
                UpdateSaveLocationMenu();
                UpdateIcon();
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

        static List<HistoryFileInfo> GetHistoryFiles()
        {
            var result = new List<HistoryFileInfo>();

            if (!Directory.Exists(SaveDirectory))
                return result;

            foreach (var file in Directory.GetFiles(SaveDirectory, "EW_paths_*.txt"))
            {
                try
                {
                    var lines = File.ReadAllLines(file)
                                    .Where(l => !string.IsNullOrWhiteSpace(l))
                                    .ToList();

                    if (lines.Count == 0)
                        continue;

                    result.Add(new HistoryFileInfo
                    {
                        FilePath = file,
                        Time = File.GetLastWriteTime(file),
                        Count = lines.Count
                    });
                }
                catch { }
            }

            return result
                .OrderByDescending(f => f.Time)
                .ToList();
        }

        static Size MeasureToolTipSize(string text)
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
            {
                var font = SystemFonts.DefaultFont;
                int maxWidth = 600;

                Size proposed = new Size(maxWidth, int.MaxValue);
                return TextRenderer.MeasureText(
                    g,
                    text,
                    font,
                    proposed,
                    TextFormatFlags.WordBreak
                );
            }
        }



        static void BuildContextMenu()
        {
            contextMenu.Items.Clear();

            LoadLastPaths();

            if (lastPaths.Count > 0)
            {
                foreach (var path in lastPaths)
                {
                    string folderName = GetFolderNameOrDrive(path);

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

            countdownMenuItem = new ToolStripMenuItem($"Next autosave will be in {countdown}s")
            {
                Enabled = true,
                ForeColor = Color.Gray,
                CheckOnClick = true,
                Checked = showProgressBar,
                ToolTipText = "Click to enable/disable progress bar"
            };

            // Клик переключает флаг
            countdownMenuItem.Click += (s, e) =>
            {
                showProgressBar = countdownMenuItem.Checked;
                UpdateIcon();
            };

            contextMenu.Items.Add(countdownMenuItem);


            intervalMenu.DropDownItems.Clear();

            AddIntervalPreset("30 sec", 30);
            AddIntervalPreset("1 min", 60);
            AddIntervalPreset("2 min", 120);
            AddIntervalPreset("5 min", 300);
            AddIntervalPreset("10 min", 600);

            if (!contextMenu.Items.Contains(intervalMenu))
            {
                contextMenu.Items.Add(intervalMenu);
            }



            contextMenu.Items.Add(new ToolStripSeparator());

            saveLocationMenu.DropDownItems.Clear();

            AddSaveLocationPreset("Temp folder", "temp");
            AddSaveLocationPreset("Program folder", "program");

            if (!contextMenu.Items.Contains(saveLocationMenu))
            {
                contextMenu.Items.Add(saveLocationMenu);
            }

            contextMenu.Items.Add(new ToolStripSeparator());

            /*
            restoreItem = new ToolStripMenuItem("Restore windows");
            restoreItem.Click += (s, e) => RestoreWindows();
            contextMenu.Items.Add(restoreItem);
            */

            var restoreMenu = new ToolStripMenuItem("Restore windows");

            var historyFiles = GetHistoryFiles();

            if (historyFiles.Count > 0)
            {
                // Restore latest
                var latest = historyFiles[0];
                var restoreLatestItem = new ToolStripMenuItem($"Restore latest ({latest.Count})");

                restoreLatestItem.Click += (s, e) => RestoreWindowsFromFile(latest.FilePath);

                restoreMenu.DropDownItems.Add(restoreLatestItem);
                restoreMenu.DropDownItems.Add(new ToolStripSeparator());

                // Остальные версии
                foreach (var hf in historyFiles)
                {
                    string text = $"{hf.Time:yyyy-MM-dd HH:mm:ss} ({hf.Count})";

                    var item = new ToolStripMenuItem(text)
                    {
                        ToolTipText = "" // важно: оставляем пустым
                    };

                    item.MouseEnter += (s, e) =>
                    {
                        string preview = BuildPreviewText(hf.FilePath);

                        Size tipSize = MeasureToolTipSize(preview);

                        Point cursorScreen = Cursor.Position;
                        Point cursorClient = contextMenu.PointToClient(cursorScreen);

                        // Показываем ВЫШЕ курсора
                        var location = new Point(
                            cursorClient.X + 10,
                            cursorClient.Y - tipSize.Height - 10
                        );

                        historyToolTip.Show(preview, contextMenu, location);
                    };



                    item.MouseLeave += (s, e) =>
                    {
                        historyToolTip.Hide(contextMenu);
                    };

                    item.Click += (s, e) =>
                        RestoreWindowsFromFile(hf.FilePath);

                    restoreMenu.DropDownItems.Add(item);
                }

            }
            else
            {
                restoreMenu.DropDownItems.Add(
                    new ToolStripMenuItem("(no saved versions)") { Enabled = false });
            }

            contextMenu.Items.Add(restoreMenu);

            //contextMenu.Items.Add("Open saved list", null, (s, e) => OpenPathsFile());
            contextMenu.Items.Add("Open saved history", null, (s, e) => OpenSaveDirectory());


            var clearItem = new ToolStripMenuItem("Clear saved history");
            clearItem.Click += (s, e) =>
            {
                try
                {
                    if (Directory.Exists(SaveDirectory))
                    {
                        foreach (var file in Directory.GetFiles(SaveDirectory, "EW_paths_*.txt"))
                            File.Delete(file);
                    }

                    lastPaths.Clear();
                    BuildContextMenu();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Clear history error: {ex}");
                }
            };
            contextMenu.Items.Add(clearItem);


            contextMenu.Items.Add("Exit", null, (s, e) => Exit()); 

            UpdateCountdown(); 
            UpdateIntervalMenu();
            UpdateSaveLocationMenu();
        }

        static Icon CreateTrayIconWithProgress(Icon baseIcon, double progress)
        {
            int width = baseIcon.Width;
            int height = baseIcon.Height;

            using (Bitmap bmp = new Bitmap(width, height))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent); // прозрачный фон

                // Рисуем базовую иконку
                using (Icon tmpIcon = (Icon)baseIcon.Clone())
                {
                    g.DrawIcon(tmpIcon, 0, 0);
                }

                // Полоска прогресса (уменьшение, выровнено по левому краю)
                int barHeight = 4;
                int barWidth = (int)(width * progress); // 1.0 = полная, 0 = пустая
                int barX = 0; // левая граница

                if (barWidth > 0)
                {
                    Rectangle barFg = new Rectangle(barX, height - barHeight, barWidth, barHeight);
                    g.FillRectangle(Brushes.LimeGreen, barFg);
                }

                IntPtr hIcon = bmp.GetHicon();
                return Icon.FromHandle(hIcon);
            }
        }


        static void RestoreWindowsFromFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return;

                var paths = File.ReadAllLines(filePath)
                                .Where(l => !string.IsNullOrWhiteSpace(l))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();

                if (paths.Count == 0)
                    return;

                var shellWindows = new ShellWindows();
                var openPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (InternetExplorer window in shellWindows)
                {
                    if (window.FullName.EndsWith("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var p = new Uri(window.LocationURL).LocalPath
                                .TrimEnd(Path.DirectorySeparatorChar);
                            openPaths.Add(p);
                        }
                        catch { }
                    }
                }

                foreach (var path in paths)
                {
                    string trimmed = path.TrimEnd(Path.DirectorySeparatorChar);
                    if (!openPaths.Contains(trimmed) && Directory.Exists(path))
                    {
                        Process.Start("explorer.exe", path);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RestoreWindowsFromFile error: {ex}");
            }
        }

        static string FormatHistoryFileName(string filePath)
        {
            var name = Path.GetFileNameWithoutExtension(filePath);

            // paths_2026-02-08_14-32-10
            if (name.StartsWith("paths_"))
            {
                var ts = name.Substring(6);
                if (DateTime.TryParseExact(
                    ts,
                    "yyyy-MM-dd_HH-mm-ss",
                    null,
                    System.Globalization.DateTimeStyles.None,
                    out var dt))
                {
                    return dt.ToString("yyyy-MM-dd HH:mm:ss");
                }
            }

            return name;
        }



        static void OpenSaveDirectory()
        {
            try
            {
                Directory.CreateDirectory(SaveDirectory);
                Process.Start("explorer.exe", SaveDirectory);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OpenSaveDirectory error: {ex}");
            }
        }


        static void AddIntervalPreset(string label, int seconds)
        {
            var item = new ToolStripMenuItem(label) { Tag = seconds };
            item.CheckOnClick = false;
            item.Click += (s, e) =>
            {
                AutoSaveIntervalSeconds = seconds;
                countdown = AutoSaveIntervalSeconds;
                UpdateIntervalMenu();
                AppSettings.SetParameter("AutoSaveIntervalSeconds", AutoSaveIntervalSeconds);
                AppSettings.Save();
            };
            intervalMenu.DropDownItems.Add(item);
        }

        static void AddSaveLocationPreset(string label, string location)
        {
            var item = new ToolStripMenuItem(label) { Tag = location };
            item.CheckOnClick = false;
            item.Click += (s, e) =>
            {
                if (location == "temp")
                {
                    SaveLocation = Path.Combine(Path.GetTempPath(), "ExplorerWatcher", "last_paths.txt");
                }
                else if (location == "program")
                {
                    SaveLocation = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last_paths.txt");
                }
                UpdateSaveLocationMenu();
                AppSettings.SetParameter("SaveLocation", location);
                AppSettings.Save();
                SaveCurrentExplorerWindows();
            };
            saveLocationMenu.DropDownItems.Add(item);
        }

        static string GetFolderNameOrDrive(string path)
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
            return folderName;
        }


        static void UpdateIntervalMenu()
        {
            foreach (ToolStripMenuItem item in intervalMenu.DropDownItems)
            {
                if (item.Tag is int seconds)
                {
                    item.Checked = (seconds == AutoSaveIntervalSeconds);
                }
            }
        }

        static void UpdateSaveLocationMenu()
        {
            foreach (ToolStripMenuItem item in saveLocationMenu.DropDownItems)
            {
                if (item.Tag is string location)
                {
                    item.Checked = (location == AppSettings.GetParameter<string>("SaveLocation"));
                }
            }
        }

        private static string FormatSeconds(uint seconds)
        {
            return seconds < 60
                ? $"{seconds} sec"
                : $"{seconds / 60} min";
        }


        static void OpenOrActivateExplorerWindow(string path)
        {
            try
            {
                var shellWindows = new ShellWindows();

                string NormalizePath(string p) => Path.GetFullPath(p).TrimEnd(Path.DirectorySeparatorChar);

                string targetPath = NormalizePath(path);

                foreach (InternetExplorer window in shellWindows)
                {
                    if (string.Equals(Path.GetFileName(window.FullName), "explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        string windowPath = "";
                        try
                        {
                            windowPath = new Uri(window.LocationURL).LocalPath;
                        }
                        catch (UriFormatException)
                        {
                            Debug.WriteLine("Failed to get window.LocationURL");
                            continue;
                        }

                        if (string.Equals(NormalizePath(windowPath), targetPath, StringComparison.OrdinalIgnoreCase))
                        {
                            IntPtr hwnd = new IntPtr(window.HWND);
                            if (IsIconic(hwnd))
                                ShowWindow(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
                            return;
                        }
                    }
                }

                // Если окно не найдено — открыть новое
                if (Directory.Exists(path))
                    Process.Start("explorer.exe", path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OpenOrActivateExplorerWindow error: {ex.Message}");
            }
        }


        static void UpdateCountdown()
        {
            if (countdownMenuItem == null)
                return;

            // Обновляем текст в меню по состоянию работы Explorer
            if (!explorerRunning)
            {
                countdownMenuItem.Text = "No Explorer windows";
            }
            else
            {
                countdownMenuItem.Text = $"Next autosave will be in {countdown}s";
            }

            if (countdownMenuItem != null)
            {
                if (!explorerRunning)
                    countdownMenuItem.Text = "No Explorer windows";
                else
                    countdownMenuItem.Text = $"Next autosave will be in {countdown}s";

                countdownMenuItem.ForeColor = Color.Gray; // сохраняем серый цвет
            }


            // Обновляем подсказку в трее
            if (trayIcon != null)
            {
                string trayText = explorerRunning
                    ? $"ExplorerWatcher: Autosave will be in {countdown}s. Double click to save"
                    : "Explorer Watcher: Double click to restore Explorer windows";

                // Ограничиваем длину текста из-за ограничений Windows
                if (trayText.Length > 63)
                    trayText = trayText.Substring(0, 63);

                trayIcon.Text = trayText;
            }

            // Обновляем иконку в трее
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

                        if (!string.IsNullOrWhiteSpace(url))
                        {
                            try
                            {
                                path = new Uri(url).LocalPath;
                            }
                            catch (UriFormatException)
                            {
                                path = "";
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                        {
                            currentPaths.Add(path);
                        }
                    }
                }

                countdown--;

                if (countdown <= 0)
                {
                    if (explorerFound)
                    {
                        var distinctCurrent = currentPaths
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();

                        // Проверка на изменения в списке путей
                        if (distinctCurrent.Count != lastPaths.Count || !distinctCurrent.SequenceEqual(lastPaths, StringComparer.OrdinalIgnoreCase))
                        {
                            lastPaths = distinctCurrent;
                            SaveLastPaths();
                        }
                    }

                    countdown = AutoSaveIntervalSeconds;
                }

                explorerRunning = explorerFound;

                UpdateIcon();
            }
            catch (Exception ex)
            {
                explorerRunning = false;
                Debug.WriteLine($"CheckExplorerWindows exception: {ex}");
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
                            Debug.WriteLine("RestoreWindows: Invalid URI in LocationURL");
                        }

                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            string trimmedPath = path.TrimEnd(Path.DirectorySeparatorChar);
                            openPaths.Add(trimmedPath);
                        }
                    }
                }

                if (lastPaths == null)
                {
                    Debug.WriteLine("RestoreWindows: lastPaths is null");
                    return;
                }

                foreach (var path in lastPaths)
                {
                    string trimmedPath = path.TrimEnd(Path.DirectorySeparatorChar);
                    if (!openPaths.Contains(trimmedPath) && Directory.Exists(path))
                    {
                        Process.Start("explorer.exe", path);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RestoreWindows exception: {ex}");
            }
        }

        /*
        static void SaveLastPaths()
        {
            try
            {
                var dir = Path.GetDirectoryName(SaveLocation);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.WriteAllLines(SaveLocation, lastPaths);
            }
            catch (Exception ex)  // Лучше ловить все исключения, не только UriFormatException
            {
                Debug.WriteLine($"SaveLastPaths error: {ex}");
            }
        }*/
        static void SaveLastPaths()
        {
            try
            {
                if (lastPaths == null || lastPaths.Count == 0)
                    return; // не сохраняем пустоту

                Directory.CreateDirectory(SaveDirectory);

                // Проверка: изменилось ли содержимое
                var lastFile = Directory.GetFiles(SaveDirectory, "EW_paths_*.txt")
                                        .OrderByDescending(File.GetLastWriteTime)
                                        .FirstOrDefault();

                if (lastFile != null)
                {
                    var previous = File.ReadAllLines(lastFile)
                                       .Where(l => !string.IsNullOrWhiteSpace(l))
                                       .ToList();

                    if (previous.SequenceEqual(lastPaths, StringComparer.OrdinalIgnoreCase))
                        return; // ничего не изменилось
                }

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string filePath = Path.Combine(SaveDirectory, $"EW_paths_{timestamp}.txt");

                File.WriteAllLines(filePath, lastPaths);

                CleanupOldHistoryFiles();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SaveLastPaths(history) error: {ex}");
            }
        }

        static void CleanupOldHistoryFiles()
        {
            try
            {
                var files = Directory.GetFiles(SaveDirectory, "EW_paths_*.txt")
                                     .OrderByDescending(File.GetLastWriteTime)
                                     .Skip(MaxHistoryFiles);

                foreach (var file in files)
                    File.Delete(file);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CleanupOldHistoryFiles error: {ex}");
            }
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
                        string path = TryGetPathFromWindow(window);
                        if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                            currentPaths.Add(path);
                    }
                }

                currentPaths = currentPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                bool pathsChanged = !currentPaths.SequenceEqual(lastSavedPaths ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);

                if (pathsChanged)
                {
                    lastSavedPaths = currentPaths;
                    lastPaths = new List<string>(lastSavedPaths);
                    SaveLastPaths();

                    var cms = trayIcon?.ContextMenuStrip;
                    if (cms != null)
                    {
                        if (cms.InvokeRequired)
                            cms.Invoke((MethodInvoker)(() =>
                            {
                                BuildContextMenu();
                                UpdateIcon();
                            }));
                        else
                        {
                            BuildContextMenu();
                            UpdateIcon();
                        }
                    }
                }

                countdown = AutoSaveIntervalSeconds;
            }
            catch (Exception ex)
            {
                explorerRunning = false;
                countdown = 0;
                Debug.WriteLine($"UpdateExplorerWindows exception: {ex}");

                var cms = trayIcon?.ContextMenuStrip;
                if (cms != null)
                {
                    if (cms.InvokeRequired)
                        cms.Invoke((MethodInvoker)(() =>
                        {
                            BuildContextMenu();
                            UpdateIcon();
                        }));
                    else
                    {
                        BuildContextMenu();
                        UpdateIcon();
                    }
                }
            }
        }

        static string TryGetPathFromWindow(InternetExplorer window)
        {
            try
            {
                return new Uri(window.LocationURL).LocalPath;
            }
            catch (UriFormatException)
            {
                Debug.WriteLine("Invalid window path");
                return "";
            }
        }

        /*
        static void LoadLastPaths()
        {
            try
            {
                if (File.Exists(SaveLocation))
                {
                    lastPaths = File.ReadAllLines(SaveLocation)
                                    .Where(line => !string.IsNullOrWhiteSpace(line))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .ToList();
                }
                else
                {
                    lastPaths = new List<string>();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadLastPaths failed: {ex}");
                lastPaths = new List<string>();
            }
        }*/

        static void LoadLastPaths()
        {
            try
            {
                if (!Directory.Exists(SaveDirectory))
                {
                    lastPaths = new List<string>();
                    return;
                }

                var lastFile = Directory.GetFiles(SaveDirectory, "EW_paths_*.txt")
                                        .OrderByDescending(File.GetLastWriteTime)
                                        .FirstOrDefault();

                if (lastFile == null)
                {
                    lastPaths = new List<string>();
                    return;
                }

                lastPaths = File.ReadAllLines(lastFile)
                                .Where(l => !string.IsNullOrWhiteSpace(l))
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadLastPaths error: {ex}");
                lastPaths = new List<string>();
            }
        }


        /*
        static void UpdateIcon()
        {
            trayIcon.Icon = explorerRunning ? iconActive : iconInactive;
        }*/
        static void UpdateIcon()
        {
            Icon baseIcon = explorerRunning ? iconActive : iconInactive;

            if (showProgressBar)
            {
                // Рассчитываем прогресс (убавление полоски справа налево)
                double progress = explorerRunning
                    ? (double)countdown / AutoSaveIntervalSeconds
                    : 0.0;

                Icon newIcon = CreateTrayIconWithProgress(baseIcon, progress);

                Icon oldIcon = trayIcon.Icon;
                trayIcon.Icon = newIcon;

                if (currentDynamicIcon != null)
                    currentDynamicIcon.Dispose();
                currentDynamicIcon = newIcon;
            }
            else
            {
                // Просто базовая иконка
                Icon oldIcon = trayIcon.Icon;
                trayIcon.Icon = baseIcon;

                if (currentDynamicIcon != null)
                {
                    currentDynamicIcon.Dispose();
                    currentDynamicIcon = null;
                }
            }

            // Обновляем текст подсказки
            if (trayIcon != null)
            {
                string trayText = explorerRunning
                    ? $"ExplorerWatcher: Autosave will be in {countdown}s. Double click to save"
                    : "Explorer Watcher: Double click to restore Explorer windows";

                if (trayText.Length > 63)
                    trayText = trayText.Substring(0, 63);

                trayIcon.Text = trayText;
            }
        }



        static void OpenPathsFile()
        {
            try
            {
                if (!File.Exists(SaveLocation))
                    File.WriteAllLines(SaveLocation, new string[0]);

                Process.Start("notepad.exe", SaveLocation);
            }
            catch (UriFormatException) 
            { 
                Debug.WriteLine("Path is empty\r\n"); 
            }
        }

        static void Exit()
        {
            trayIcon.Visible = false;
            Application.Exit();
        }


        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_RESTORE = 9;

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
