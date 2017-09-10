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
                //IsHighlightCellLines = true,
            };
            Svo.AddSection(document);

            var fileName = document.SavePng(0, "Temp.png", 120);
            //Process.Start("LiveViewer", fileName);
            
            //Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
        }
    }
}
