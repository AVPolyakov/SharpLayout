using System;
using System.Collections.Generic;
using System.IO;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public class Document
    {
        public static bool CollectCallerInfo;

        public List<Func<SectionGroup>> SectionGroups { get; } = new List<Func<SectionGroup>>();

        public Section Add(Section section)
        {
            return Add(new SectionGroup()).Add(section);
        }

        public SectionGroup Add(SectionGroup sectionGroup)
        {
            Add(() => sectionGroup);
            return sectionGroup;
        }

        public void Add(Func<SectionGroup> sectionGroup)
        {
            SectionGroups.Add(sectionGroup);
        }

        public bool CellsAreHighlighted { get; set; }

        public bool R1C1AreVisible { get; set; }

        public bool ParagraphsAreHighlighted { get; set; }

        public bool CellLineNumbersAreVisible { get; set; }

	    public bool ExpressionVisible { get; set; }

	    public byte[] CreatePdf()
        {
            using (var pdfDocument = new PdfDocument())
            {
	            Render(pdfDocument);
	            using (var stream = new MemoryStream())
                {
                    pdfDocument.Save(stream);
                    return stream.ToArray();
                }
            }
        }

	    public void Render(PdfDocument pdfDocument)
	    {
		    pdfDocument.ViewerPreferences.Elements.SetName("/PrintScaling", "/None");
		    pdfDocument.Info.Creator = "SharpLayout";
            foreach (var sectionGroupFunc in SectionGroups)
            {
                var sectionGroup = sectionGroupFunc();
                if (sectionGroup.Sections.Count == 0) continue;
                var page = pdfDocument.AddPage();
                page.Size = PageSize.A4;
                page.Orientation = sectionGroup.Sections[0].PageSettings.Orientation;
                using (var xGraphics = XGraphics.FromPdfPage(page))
                {
                    TableRenderer.Draw(xGraphics, (pageIndex, action, section) => {
                        var addPage = pdfDocument.AddPage();
                        addPage.Size = PageSize.A4;
                        addPage.Orientation = section.PageSettings.Orientation;
                        using (var xGraphics2 = XGraphics.FromPdfPage(addPage))
                            action(xGraphics2);
                    }, this, GraphicsType.Pdf, sectionGroup);
                }
            }
	    }

        public string SavePdf(string path)
        {
            File.WriteAllBytes(path, CreatePdf());
            return path;
        }

        public Tuple<List<byte[]>> CreatePng()
        {
            throw new NotImplementedException();
        }
    }
}