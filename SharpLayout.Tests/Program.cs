using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static SharpLayout.Tests.Tests;

namespace SharpLayout.Tests
{
	class Program
	{
	    static void Main()
	    {
	        PageSettings pageSettings;
	        List<Table> tables;
	        Svo.GetContent(out pageSettings, out tables, true);
            Process.Start(CreatePdf(pageSettings, tables));
	        //Process.Start(SavePng(pageSettings, tables, 0));
	    }

	    public static string CreatePdf(PageSettings pageSettings, IEnumerable<Table> tables)
	    {
	        string filename;
	        using (var pdfDocument = new PdfDocument())
	        {
	            pdfDocument.ViewerPreferences.Elements.SetName("/PrintScaling", "/None");
	            var page = pdfDocument.AddPage();
	            page.Orientation = pageSettings.Orientation;
	            using (var xGraphics = XGraphics.FromPdfPage(page))
	                TableRenderer.Draw(xGraphics, pageSettings, (pageIndex, action) => {
	                    var addPage = pdfDocument.AddPage();
                        addPage.Orientation = pageSettings.Orientation;
	                    using (var xGraphics2 = XGraphics.FromPdfPage(addPage))
	                        action(xGraphics2);
	                }, tables);
	            filename = $"HelloWorld_tempfile{Guid.NewGuid():N}.pdf";
	            pdfDocument.Save(filename);
	        }
	        return filename;
	    }

	    public static string SavePng(PageSettings pageSettings, IEnumerable<Table> tables, int pageNumber)
	    {
	        const string filename = "temp.png";
            File.WriteAllBytes(filename, CreatePng(pageSettings, tables)[pageNumber]);
	        return filename;
	    }
	}
}
