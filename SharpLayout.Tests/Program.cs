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
	        PaymentOrder.GetContent(out PageSettings pageSettings, out List<Table> tables);
            Process.Start(CreatePdf(pageSettings, tables));
	        //Process.Start(SavePng(pageSettings, tables, 0));
	    }

	    public static string CreatePdf(PageSettings pageSettings, IEnumerable<Table> tables)
	    {
	        string filename;
	        using (var pdfDocument = new PdfDocument())
	        {
	            pdfDocument.ViewerPreferences.Elements.SetName("/PrintScaling", "/None");
	            using (var xGraphics = XGraphics.FromPdfPage(pdfDocument.AddPage()))
	                TableRenderer.Draw(xGraphics, pageSettings, (pageIndex, action) =>  {
	                    using (var xGraphics2 = XGraphics.FromPdfPage(pdfDocument.AddPage()))
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
