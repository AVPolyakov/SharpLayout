using System;
using System.Diagnostics;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
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
            if (alwaysShowWindow || Process.GetProcessesByName("LiveViewer").Length <= 0)
            {
                string arguments;
                if (findId)
                {
                    const string solutionName = "SharpLayout";
                    var firstOrDefault = Process.GetProcesses().FirstOrDefault(p => p.ProcessName == "devenv" &&
                        p.MainWindowTitle.Contains(solutionName));
                    if (firstOrDefault != null)
                        arguments = $"{fileName} {firstOrDefault.Id}";
                    else
                        arguments = fileName;
                }
                else
                    arguments = fileName;
                Process.Start("LiveViewer", arguments);
            }
        }
    }
}
