﻿using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Newtonsoft.Json;
using SharpLayout;
using static System.Math;
using Process = System.Diagnostics.Process;

namespace LiveViewer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            pictureBox.MouseEnter += (sender, args) => panel1.Focus();
            panel1.KeyDown += Panel1OnKeyDown;
            ApplySettings();
            LoadFile(Environment.GetCommandLineArgs());
            switch (Ide)
            {
                case Ide.VisualStudio:
                    dte2 = GetDTE();
                    MessageFilter.Register();
                    break;
                case Ide.Rider:
                    Text = $"{Text} (Rider)";
                    break;
            }
        }

        private Ide Ide
        {
            get
            {
                if (_lineArgs.Length >= 3 && _lineArgs[2] == "rider") 
                    return Ide.Rider;
                return Ide.VisualStudio;
            }
        }

        private DTE2 GetDTE()
        {
            if (_lineArgs.Length >= 4)
            {
                var processId = int.Parse(_lineArgs[3]);
                DTE2 result;
                try
                {
                    result = DTEFinder.GetDTE(processId);
                }
                catch
                {
                    ShowRunningVisualStudioIsRequired();
                    return null;
                }
                if (result == null)
                    MessageBox.Show($"Visual Studio 2019 with process id {processId} not found.",
                        "LiveViewer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    var mainWindowTitle = Process.GetProcessById(processId).MainWindowTitle?.Trim();
                    Text = $"{Text} ({mainWindowTitle})";
                }
                return result;
            }
            else
                try
                {
                    return DTEFinder.GetDTE(processId: null);
                }
                catch
                {
                    ShowRunningVisualStudioIsRequired();
                    return null;
                }
        }

        private static void ShowRunningVisualStudioIsRequired()
        {
            MessageBox.Show("А running Visual Studio 2019 is required for fully-featured operation of Live Viewer.",
                "LiveViewer", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ApplySettings()
        {
            ResizeEnd += (sender, args) => {
                if (WindowState == FormWindowState.Normal)
                    File.WriteAllText(SettingsPath,
                        JsonConvert.SerializeObject(new Settings {
                            Location = Location,
                            Size = Size,
                            WindowState = WindowState
                        }, Formatting.Indented));
            };
            FormClosing += (sender, args) => {
                Point? location;
                Size? size;
                if (WindowState == FormWindowState.Normal)
                {
                    location = Location;
                    size = Size;
                }
                else
                {
                    if (File.Exists(SettingsPath))
                    {
                        var settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(SettingsPath));                        
                        location = settings.Location;
                        size = settings.Size;
                    }
                    else
                    {
                        location = null;
                        size = null;
                    }
                }
                File.WriteAllText(SettingsPath,
                    JsonConvert.SerializeObject(new Settings {
                        Location = location,
                        Size = size,
                        WindowState = WindowState
                    }, Formatting.Indented));
            };
            Load += (sender, args) => {
                if (File.Exists(SettingsPath))
                {
                    var settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(SettingsPath));
                    if (settings.Location.HasValue) Location = settings.Location.Value;
                    if (settings.Size.HasValue) Size = settings.Size.Value;
                    WindowState = settings.WindowState;
                }
            };
        }

        private static string SettingsPath => Path.ChangeExtension(typeof(Program).Assembly.Location, ".json");

        private string[] _lineArgs;
        private string _imageLocation;
        public void LoadFile(string[] lineArgs)
        {
            _lineArgs = lineArgs;
            if (_lineArgs.Length >= 2)
            {
                _imageLocation = Path.GetFullPath(_lineArgs[1]);
                LoadImage();
                fileSystemWatcher.Path = Path.GetDirectoryName(_imageLocation);
            }
        }

        private void LoadImage()
        {
            Image image;
            try
            {
                using (var stream = File.OpenRead(_imageLocation)) 
                    image = Image.FromStream(stream);
            }
            catch
            {
                return;
            }
            pictureBox.Image = image;
        }

        private void fileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (string.Equals(e.FullPath, _imageLocation, StringComparison.InvariantCultureIgnoreCase)) 
                LoadImage();
        }

        private readonly DTE2 dte2;

        private void pictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (ModifierKeys != Keys.Control) return;
            var bitmapInfo = GetSyncBitmapInfo();
            var firstOrDefault = bitmapInfo.PageInfo.ItemInfos.Where(info => {
                    var x = info.X.ToPixel(bitmapInfo);
                    var y = info.Y.ToPixel(bitmapInfo);
                    var width = info.Width.ToPixel(bitmapInfo);
                    var heigth = info.Height.ToPixel(bitmapInfo);
                    return x <= e.X && e.X <= x + width &&
                        y <= e.Y && e.Y <= y + heigth &&
                        info.CallerInfos.Count > 0;
                })
                .OrderByDescending(_ => _.TableLevel).ThenByDescending(_ => _.Level)
                .FirstOrDefault();
            if (firstOrDefault != null)
            {
                multipleLineMenuStrip.Items.Clear();
                var callerInfos = firstOrDefault.CallerInfos.Select(_ => new {_.Line, _.FilePath}).Distinct().ToList();
                if (callerInfos.Count == 1)
                    GotoLine(callerInfos[0].Line, callerInfos[0].FilePath);
                else
                    foreach (var callerInfo in callerInfos)
                    {
                        var text = File.ReadAllLines(callerInfo.FilePath)[callerInfo.Line - 1].Trim();
                        multipleLineMenuStrip.Items.Add($"{callerInfo.Line} {text}", null,
                            (o, args) => GotoLine(callerInfo.Line, callerInfo.FilePath));
                    }
                multipleLineMenuStrip.Show(pictureBox, new Point(e.X, e.Y));
            }
            else
                MessageBox.Show(this, "В этой части картинки нет ссылки на исходный код.",
                    "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GotoLine(int line, string filePath)
        {
            switch (Ide)
            {
                case Ide.VisualStudio:
                    dte2.ItemOperations.OpenFile(filePath, Constants.vsViewKindTextView);
                    var textSelection = (TextSelection) dte2.ActiveDocument.Selection;
                    textSelection.GotoLine(line);
                    textSelection.WordRight();
                    SetForegroundWindow(new IntPtr(dte2.MainWindow.HWnd));
                    break;
                case Ide.Rider:
                    var processes = Process.GetProcessesByName("rider64");
                    var mainProcess = processes.FirstOrDefault(_ => !string.IsNullOrEmpty(_.MainWindowTitle));
                    var process= mainProcess ?? processes[0];
                    Process.Start(process.MainModule.FileName, $"--line {line} {filePath}");
                    if (mainProcess != null) 
                        SetForegroundWindow(mainProcess.MainWindowHandle);
                    break;
            }
        }

        private SyncBitmapInfo GetSyncBitmapInfo()
        {
            return JsonConvert.DeserializeObject<SyncBitmapInfo>(
                File.ReadAllText(Path.ChangeExtension(_imageLocation, ".json")));
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private void Panel1OnKeyDown(object sender, KeyEventArgs e)
        {
            int? value;
            switch (e.KeyData)
            {
                case Keys.W:
                    value = selectionWidth;
                    break;
                case Keys.H:
                    value = selectionHeight;
                    break;
                default:
                    value = null;
                    break;
            }
            if (value.HasValue)
            {
                var text = ((int) Round(1.0 * value.Value * 254 / GetSyncBitmapInfo().Resolution,
                    MidpointRounding.AwayFromZero)).ToString();
                switch (Ide)
                {
                    case Ide.VisualStudio:
                        var textSelection = (TextSelection) dte2.ActiveDocument.Selection;
                        textSelection.Insert(text);
                        textSelection.CharLeft(true, text.Length);
                        SetForegroundWindow(new IntPtr(dte2.MainWindow.HWnd));
                        dte2.ExecuteCommand("Debug.StartWithoutDebugging");
                        break;
                    case Ide.Rider:
                        Clipboard.SetText(text);
                        break;
                }
            }
        }

        private int selectionX;
        private int selectionY;
        private int selectionWidth;
        private int selectionHeight;

        private void pictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                selectionX = e.X;
                selectionY = e.Y;
            }
            pictureBox.Refresh();
            selectionWidth = 0;
            selectionHeight = 0;
        }

        private void pictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                pictureBox.Refresh();
                selectionWidth = Abs(e.X - selectionX);
                selectionHeight = Abs(e.Y - selectionY);
                pictureBox.CreateGraphics().DrawRectangle(new Pen(Color.Black, 1) { DashStyle = DashStyle.DashDot },
                    Min(e.X, selectionX),
                    Min(e.Y, selectionY), 
                    selectionWidth, 
                    selectionHeight);
            }
        }
    }

    internal enum Ide
    {
        VisualStudio,
        Rider
    }
}
