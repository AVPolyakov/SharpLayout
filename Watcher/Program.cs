using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using OtherAssembly1;
using OtherAssembly2;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using SharpLayout;

namespace Watcher
{
    static class Program
    {
        public static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var settingsPath = WatcherSettingsProvider.FilePath;
            var outputPath = @"..\Starter\bin\Debug\netcoreapp3.0\Temp.png";
            var fullOutputPath = Path.Combine(Path.GetDirectoryName(settingsPath), outputPath);
            
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add("Starting...", new Font("Consolas", 9.5, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode)))));
            document.SavePng(0, fullOutputPath, 120);
            
            fullOutputPath.StartLiveViewer(true);
            
            SharpLayout.WatcherCore.Watcher.Start(
                settingsPath: settingsPath,
                assemblies: new[] {
                    typeof(IA),
                    typeof(A),
                }.Select(_ => _.Assembly).ToArray(),
                parameterFunc: info => {
                    if (info.ParameterType == typeof(IA))
                        return typeof(A);
                    throw new Exception($"Factory is not specified for type {info.ParameterType}");
                },
                outputPath: outputPath);
        }
        
        public static void StartLiveViewer(this string fileName, bool alwaysShowWindow, bool findId = true)
        {
            var processes = Process.GetProcessesByName("LiveViewer");
            if (processes.Length <= 0)
            {
                string arguments;
                if (findId && Ide == vs)
                {
                    const string solutionName = "SharpLayout";
                    var firstOrDefault = Process.GetProcesses().FirstOrDefault(p => p.ProcessName == "devenv" &&
                        p.MainWindowTitle.Contains(solutionName));
                    if (firstOrDefault != null)
                        arguments = $" {firstOrDefault.Id}";
                    else
                        arguments = "";
                }
                else
                    arguments = "";
                Process.Start("LiveViewer", fileName + " " + Ide + arguments);
            }
            else
            {
                if (alwaysShowWindow)
                    SetForegroundWindow(processes[0].MainWindowHandle);
            }
        }

        private static string Ide
        {
            get
            {
                switch (Environment.UserName)
                {
                    case "APolyakov":
                        return "rider";
                    default:
                        return vs;
                }
            }
        }

        private const string vs = "vs";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
