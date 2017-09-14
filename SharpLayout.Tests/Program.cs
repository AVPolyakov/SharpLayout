using System;
using System.Diagnostics;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            var document = new Document {
                IsHighlightCells = true,
                IsHighlightParagraphs = true,
                //ShowCellLine = true
            };
            PaymentOrder.AddSection(document);

            document.SavePng(0, "Temp.png", 120).StartLiveViewer(false);

            //Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
        }

        private static void StartLiveViewer(this string fileName, bool alwaysShowWindow)
        {
            if (alwaysShowWindow || Process.GetProcessesByName("LiveViewer").Length <= 0)
                Process.Start("LiveViewer", fileName);
        }
    }
}
