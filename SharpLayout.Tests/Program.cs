using System;
using System.Diagnostics;

namespace SharpLayout.Tests
{
    class Program
    {
        static void Main()
        {
            var document = new Document {IsHighlightCellLine = true};
            PaymentOrder.AddSection(document);
            Svo.AddSection(document);
            Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            //Process.Start(document.SavePng(0, "temp.png"));
        }
    }
}
