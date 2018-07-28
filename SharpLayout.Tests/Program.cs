using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using PdfSharp;
using PdfSharp.Drawing;
using static SharpLayout.Direction;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    static class Program
    {
        static async Task Main()
        {
            M1(50);
            for (var j = 1; j <= 17; j++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                await Task.WhenAll(Enumerable.Range(1, j).Select(i => Task.Run(() => {
                    //
                    M1(3000);
                })));
                stopwatch.Stop();
                Console.WriteLine($"{j} {stopwatch.ElapsedMilliseconds}");
            }

            //var path = $"Temp_{Guid.NewGuid():N}.pdf";
            //File.WriteAllBytes(path, bytes);
            //Process.Start(path);
        }

        private static byte[] M1(int count)
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1),
                Orientation = PageOrientation.Landscape
            }));
            var table = section.AddTable().Border(Styles.BorderWidth).Font(Styles.TimesNewRoman8);
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
                            .Add("QWEQWE QWETQWTQWRTWERT QWETQWET QWEQWR ASDASF"));
                    }
                });
            }

            var bytes = document.CreatePdf();
            return bytes;
        }

        private static void ShowException(Exception e)
        {
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add($"{e}", new XFont("Consolas", 9.5))));
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
