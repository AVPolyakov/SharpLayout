using System;
using System.Diagnostics;

namespace SharpLayout.Tests
{
    class Program
    {
        static void Main()
        {
            var document = new Document {
                IsHighlightCells = true,
                IsHighlightParagraphs = true,
                IsHighlightCellLines = true,
            };
            PaymentOrder.AddSection(document);
            //Svo.AddSection(document);
            document.SavePng(0, "temp2.png", 120);
            //document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf");
            //document.SavePng(0, "temp.png");
            Process.Start(@"C:\Users\APolyakov\Documents\Visual Studio 2015\Projects\WindowsFormsApplication2\WindowsFormsApplication2\bin\Debug\WindowsFormsApplication2.exe");
        }
    }
}
