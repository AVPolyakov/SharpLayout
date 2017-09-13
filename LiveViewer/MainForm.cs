using System;
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

namespace LiveViewer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            ApplySettings();
            LoadFile(Environment.GetCommandLineArgs());
        }

        private void ApplySettings()
        {
            FormClosing += (sender, args) => {
                File.WriteAllText(SettingsPath,
                    JsonConvert.SerializeObject(new Settings {
                        WindowState = WindowState,
                        Location = Location,
                        Size = Size
                    }, Formatting.Indented));
            };
            Load += (sender, args) => {
                if (File.Exists(SettingsPath))
                {
                    var settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(SettingsPath));
                    WindowState = settings.WindowState;
                    Location = settings.Location;
                    Size = settings.Size;
                }
            };
        }

        private static string SettingsPath => Path.ChangeExtension(typeof(Program).Assembly.Location, ".json");

        public void LoadFile(string[] lineArgs)
        {
            if (lineArgs.Length >= 2)
            {
                var imageLocation = Path.GetFullPath(lineArgs[1]);
                pictureBox.ImageLocation = imageLocation;
                fileSystemWatcher.Path = Path.GetDirectoryName(imageLocation);
            }
        }

        private void fileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            if (string.Equals(e.FullPath, pictureBox.ImageLocation, StringComparison.InvariantCultureIgnoreCase))
                pictureBox.ImageLocation = pictureBox.ImageLocation;
        }

        private readonly DTE2 dte2 = (DTE2)Marshal.GetActiveObject("VisualStudio.DTE.14.0");

        private void pictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (ModifierKeys != Keys.Control) return;
            var bitmapInfo = GetSyncBitmapInfo();
            var firstOrDefault = bitmapInfo.PageInfo.CellInfos.Where(info => {
                    var x = info.X.ToPixel(bitmapInfo);
                    var y = info.Y.ToPixel(bitmapInfo);
                    var width = info.Width.ToPixel(bitmapInfo);
                    var heigth = info.Height.ToPixel(bitmapInfo);
                    return x <= e.X && e.X <= x + width &&
                        y <= e.Y && e.Y <= y + heigth &&
                        info.Line.HasValue && info.FilePath != null;
                })
                .OrderByDescending(_ => _.TableLevel)
                .FirstOrDefault();
            if (firstOrDefault != null)
            {
                dte2.ItemOperations.OpenFile(
                    firstOrDefault.FilePath,
                    Constants.vsViewKindTextView);
                var textSelection = (TextSelection)dte2.ActiveDocument.Selection;
                textSelection.GotoLine(firstOrDefault.Line.Value);
                textSelection.WordRight();
                SetForegroundWindow(new IntPtr(dte2.MainWindow.HWnd));
            }
            else
                MessageBox.Show(this, "В этой части картинки нет ссылки на исходный код.",
                    "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private SyncBitmapInfo GetSyncBitmapInfo()
        {
            return JsonConvert.DeserializeObject<SyncBitmapInfo>(
                File.ReadAllText(Path.ChangeExtension(pictureBox.ImageLocation, ".json")));
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
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
                var textSelection = (TextSelection) dte2.ActiveDocument.Selection;
                var text = ((int) Round(1.0 * value.Value * 254 / GetSyncBitmapInfo().Resolution,
                    MidpointRounding.AwayFromZero)).ToString();
                textSelection.Insert(text);
                textSelection.CharLeft(true, text.Length);
                SetForegroundWindow(new IntPtr(dte2.MainWindow.HWnd));
                dte2.ExecuteCommand("Debug.StartWithoutDebugging");
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
}
