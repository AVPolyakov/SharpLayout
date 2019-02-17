using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public class Document
    {
        public static bool CollectCallerInfo;

        public List<Func<Section>> Sections { get; } = new List<Func<Section>>();

        public Section Add(Section section)
        {
            Add(() => section);
            return section;
        }

        public void Add(Func<Section> section)
        {
            Sections.Add(section);
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
            if (Sections.Count == 0) return;
            foreach (var sectionFunc in Sections)
            {
                var section = sectionFunc();
                var page = pdfDocument.AddPage();
                page.Size = PageSize.A4;
                page.Orientation = section.PageSettings.Orientation;
                using (var xGraphics = XGraphics.FromPdfPage(page))
                {
                    TableRenderer.Draw(xGraphics, (pageIndex, action) => {
                        var addPage = pdfDocument.AddPage();
                        addPage.Size = PageSize.A4;
                        addPage.Orientation = section.PageSettings.Orientation;
                        using (var xGraphics2 = XGraphics.FromPdfPage(addPage))
                            action(xGraphics2);
                    }, this, GraphicsType.Pdf, section);
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
            return CreateImage(ImageFormat.Png, resolution);
        }

        public Tuple<List<byte[]>, List<SyncBitmapInfo>> CreateImage(ImageFormat imageFormat, int resolution = defaultResolution)
        {
            var list = new List<byte[]>();
            var syncBitmapInfos = new List<SyncBitmapInfo>();
            foreach (var sectionFunc in Sections)
            {
                var section = sectionFunc();
                var pages = new List<byte[]> {null};
                var syncPageInfos = FillBitmap(xGraphics => TableRenderer.Draw(xGraphics,
                        (pageIndex, action) => {
                            FillBitmap(graphics => {
                                    action(graphics);
                                    return new { };
                                },
                                bitmap => pages.Add(ToBytes(bitmap, imageFormat)),
                                section.PageSettings, resolution);
                        }, this, GraphicsType.Image, section),
                    bitmap => pages[0] = ToBytes(bitmap, imageFormat),
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

        public static byte[] ToBytes(Bitmap bitmap, ImageFormat imageFormat)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, imageFormat);
                return stream.ToArray();
            }
        }
    }
}