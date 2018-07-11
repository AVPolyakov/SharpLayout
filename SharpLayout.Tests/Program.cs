using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using PdfSharp.Drawing;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            Document.CollectCallerInfo = true;
            try
            {
                var document = new Document {
                    //CellsAreHighlighted = true,
                    //R1C1AreVisible = true,
                    //ParagraphsAreHighlighted = true,
                    //CellLineNumbersAreVisible = true,
                    //ExpressionVisible = true,
                };
                PaymentOrder.AddSection(document, new PaymentData {IncomingDate = DateTime.Now, OutcomingDate = DateTime.Now});
                //Svo.AddSection(document);
                //ContractDealPassport.AddSection(document);
                //LoanAgreementDealPassport.AddSection(document);

                document.SavePng(0, "Temp.png", 120).StartLiveViewer(true);

                //Process.Start(document.SavePng(0, "Temp2.png")); //open with Paint.NET
                //Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            }
            catch (Exception e)
            {
                ShowException(e);
            }
        }

        private static void ShowException(Exception e)
        {
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
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
