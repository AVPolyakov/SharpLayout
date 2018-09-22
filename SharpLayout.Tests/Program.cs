using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using PdfSharp;
using PdfSharp.Drawing;
using static SharpLayout.Direction;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            for (var j = 1; j <= 5; j++)
                M1(3000);
            for (var j = 1; j <= 20; j++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                M1(3000);
                stopwatch.Stop();
                Console.WriteLine($"{stopwatch.ElapsedMilliseconds}");
            }
        }

        private static byte[] M1(int count)
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings
            {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1),
                Orientation = PageOrientation.Landscape
            }));
            var table = section.AddTable().Border(Styles.BorderWidth);
            var columnCount = 14;
            for (var i = 0; i < columnCount; i++)
                table.AddColumn();
            table.Columns.ToArray().Distribute(section.PageSettings.PageWidthWithoutMargins);
            for (var j = 0; j < count; j++)
            {
                table.AddRow(r => {
                    for (var i = 0; i < columnCount; i++)
                    {
                        r[table.Columns[i]].Add(new Paragraph().Margin(Left | Right, Cm(0.05))
                            .Add("QWEQWE QWETQWTQWRTWERT QWETQWET QWEQWR ASDASF", Styles.TimesNewRoman8));
                    }
                });
            }

            var bytes = document.CreatePdf();

            //var path = $"Temp_{Guid.NewGuid():N}.pdf";
            //File.WriteAllBytes(path, bytes);
            //Process.Start(path);

            return bytes;
        }

        private static void ShowException(Exception e)
        {
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add($"{e}", new Font("Consolas", 9.5, XFontStyle.Regular, Styles.PdfOptions))));
            document.SavePng(0, "Temp.png", 120).StartLiveViewer(true, false);
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

        private static string Ide => vs;

        private const string vs = "vs";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
