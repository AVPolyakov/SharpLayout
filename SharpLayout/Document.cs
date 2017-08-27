using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public class Document
    {
        public List<Section> Sections { get; } = new List<Section>();

        public void Add(Section section) => Sections.Add(section);

        public bool IsHighlightCells { get; set; } = false;

        public bool IsHighlightCellLine { get; set; } = false;

        public byte[] CreatePdf()
        {
            using (var pdfDocument = new PdfDocument())
            {
                pdfDocument.ViewerPreferences.Elements.SetName("/PrintScaling", "/None");
                pdfDocument.Info.Creator = "SharpLayout";
                foreach (var section in Sections)
                {
                    var page = pdfDocument.AddPage();
                    page.Orientation = section.PageSettings.Orientation;
                    using (var xGraphics = XGraphics.FromPdfPage(page))
                        TableRenderer.Draw(xGraphics, section.PageSettings, (pageIndex, action) => {
                            var addPage = pdfDocument.AddPage();
                            addPage.Orientation = section.PageSettings.Orientation;
                            using (var xGraphics2 = XGraphics.FromPdfPage(addPage))
                                action(xGraphics2);
                        }, section.Tables, this);
                }
                using (var stream = new MemoryStream())
                {
                    pdfDocument.Save(stream);
                    return stream.ToArray();
                }
            }
        }

        public string SavePdf(string path)
        {
            File.WriteAllBytes(path, CreatePdf());
            return path;
        }

        public List<byte[]> CreatePng()
        {
            var list = new List<byte[]>();
            foreach (var section in Sections)
            {
                var pages = new List<byte[]> {null};
                FillBitmap(xGraphics => TableRenderer.Draw(xGraphics, section.PageSettings,
                        (pageIndex, action) => FillBitmap(action, bitmap => pages.Add(ToBytes(bitmap)), section.PageSettings), section.Tables, this),
                    bitmap => pages[0] = ToBytes(bitmap), section.PageSettings);
                list.AddRange(pages);
            }
            return list;
        }

        public string SavePng(int pageNumber, string tempPng)
        {
            File.WriteAllBytes(tempPng, CreatePng()[pageNumber]);
            return tempPng;
        }

        public static void FillBitmap(Action<XGraphics> action, Action<Bitmap> action2, PageSettings pageSettings)
        {
            const int resolution = 254;
            var horzPixels = (int) (new XUnit(pageSettings.PageWidth).Inch * resolution);
            var vertPixels = (int) (new XUnit(pageSettings.PageHeight).Inch * resolution);
            using (var bitmap = new Bitmap(horzPixels, vertPixels))
            {
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, horzPixels, vertPixels);
                    using (var xGraphics = XGraphics.FromGraphics(graphics, new XSize(horzPixels, vertPixels)))
                    {
                        xGraphics.ScaleTransform(resolution / 72d);
                        action(xGraphics);
                    }
                }
                bitmap.SetResolution(resolution, resolution);
                action2(bitmap);
            }
        }

        public static byte[] ToBytes(Bitmap bitmap)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }
    }
}