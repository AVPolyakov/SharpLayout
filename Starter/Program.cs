using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Examples;
using PdfSharp.Drawing;
using SharpLayout;

namespace Starter
{
    static class Program
    {
        static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Document.CollectCallerInfo = true;
            try
            {
                var document = new Document {
                    //CellsAreHighlighted = true,
                    //R1C1AreVisible = true,
                    //ParagraphsAreHighlighted = true,
                    //CellLineNumbersAreVisible = true,
                    //ExpressionVisible = true,
                    TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit
                };
                LegalEntityCreation.AddSection(document, GetData());

                var path = "Temp.png";

                File.WriteAllBytes(path, document.CreateImage(ImageFormat.Tiff, resolution: 300).Item1[0]);
                
                StartProcess(path); //open with Paint.NET

                //path.StartLiveViewer(alwaysShowWindow: true);
                //StartProcess(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            }
            catch (Exception e)
            {
                ShowException(e);
            }
        }

        private static LegalEntityCreationData GetData()
        {
            return new LegalEntityCreationData
            {
                FullName = "Полное имя aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                Name = "Сокращенное имя",
                PostalCode = "123456",
                Subject = "АБ",
                AreaName = "Большой улус аааааааааааааааааааааааааааааааа",
                Area = "Улус №1",
                City = "Город №1",
                CityName = "Большой город",
            };
        }

        public static void StartProcess(string fileName)
        {
            Process.Start(new ProcessStartInfo("cmd", $"/c start {fileName}") {CreateNoWindow = true});
        }

        private static void ShowException(Exception e)
        {
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
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
                if (findId && Ide == "vs")
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

        private static string Ide => Environment.UserName switch {
            "APolyakov" => "rider",
            _ => "vs"
        };

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
