﻿using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Newtonsoft.Json;
using SharpLayout;

namespace LiveViewer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            LoadFile(Environment.GetCommandLineArgs());
        }

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
            var bitmapInfo = JsonConvert.DeserializeObject<SyncBitmapInfo>(
                File.ReadAllText(Path.ChangeExtension(pictureBox.ImageLocation, ".json")));
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

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
