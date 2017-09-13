using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public class Document
    {
        public List<Section> Sections { get; } = new List<Section>();

        public Section Add(Section section)
        {
            Sections.Add(section);
            return section;
        }

        public bool IsHighlightCells { get; set; } = false;

        public bool IsHighlightParagraphs { get; set; } = false;

        public bool IsHighlightCellLines { get; set; } = false;

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

        public Tuple<List<byte[]>, List<SyncBitmapInfo>> CreatePng(int resolution = defaultResolution)
        {
            var list = new List<byte[]>();
            var syncBitmapInfos = new List<SyncBitmapInfo>();
            foreach (var section in Sections)
            {
                var pages = new List<byte[]> {null};
                var syncPageInfos = FillBitmap(xGraphics => TableRenderer.Draw(xGraphics, section.PageSettings,
                        (pageIndex, action) => {
                            FillBitmap(graphics => {
                                    action(graphics);
                                    return new { };
                                },
                                bitmap => pages.Add(ToBytes(bitmap)),
                                section.PageSettings, resolution);
                        }, section.Tables, this),
                    bitmap => pages[0] = ToBytes(bitmap),
                    section.PageSettings, resolution);
                syncBitmapInfos.AddRange(syncPageInfos.Select(pageInfo => new SyncBitmapInfo {
                    PageInfo = pageInfo,
                    Resolution = resolution,
                    HorizontalPixelCount = HorizontalPixelCount(section.PageSettings, resolution),
                    VerticalPixelCount = VerticalPixelCount(section.PageSettings, resolution),
                }));
                list.AddRange(pages);
            }
            return Tuple.Create(list, syncBitmapInfos);
        }

        public string SavePng(int pageNumber, string path, int resolution = defaultResolution)
        {
            var tuple = CreatePng(resolution);
            File.WriteAllBytes(path, tuple.Item1[pageNumber]);
            File.WriteAllText(Path.ChangeExtension(path, ".json"),
                JsonConvert.SerializeObject(tuple.Item2[pageNumber], Formatting.Indented));
            return path;
        }

        private const int defaultResolution = 254;

        public static T FillBitmap<T>(Func<XGraphics, T> func, Action<Bitmap> action2, PageSettings pageSettings, int resolution)
        {
            var horizontalPixelCount = HorizontalPixelCount(pageSettings, resolution);
            var verticalPixelCount = VerticalPixelCount(pageSettings, resolution);
            using (var bitmap = new Bitmap(horizontalPixelCount, verticalPixelCount))
            {
                T result;
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, horizontalPixelCount, verticalPixelCount);
                    using (var xGraphics = XGraphics.FromGraphics(graphics, new XSize(horizontalPixelCount, verticalPixelCount)))
                    {
                        xGraphics.ScaleTransform(resolution / 72d);
                        result = func(xGraphics);
                    }
                }
                bitmap.SetResolution(resolution, resolution);
                action2(bitmap);
                return result;
            }
        }

        private static int VerticalPixelCount(PageSettings pageSettings, int resolution) 
            => (int) (new XUnit(pageSettings.PageHeight).Inch * resolution);

        private static int HorizontalPixelCount(PageSettings pageSettings, int resolution) 
            => (int) (new XUnit(pageSettings.PageWidth).Inch * resolution);

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