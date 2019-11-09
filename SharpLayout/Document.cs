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
                    TableRenderer.Draw(new Graphics(xGraphics), (pageIndex, action, section) => {
                        var addPage = pdfDocument.AddPage();
                        addPage.Size = PageSize.A4;
                        addPage.Orientation = section.PageSettings.Orientation;
                        using (var xGraphics2 = XGraphics.FromPdfPage(addPage))
                            action(new Graphics(xGraphics2));
                    }, this, GraphicsType.Pdf, sectionGroup, AllPageFilter.Instance);
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
            return CreatePng(resolution, AllPageFilter.Instance);
        }
        
        private Tuple<List<byte[]>, List<SyncBitmapInfo>> CreatePng(int resolution, IPageFilter pageFilter)
        {
            return CreateImage(ImageFormat.Png, pageFilter, resolution);
        }

        public Tuple<List<byte[]>, List<SyncBitmapInfo>> CreateImage(ImageFormat imageFormat, int resolution = defaultResolution)
        {
            return CreateImage(imageFormat, AllPageFilter.Instance, resolution);
        }

        private Tuple<List<byte[]>, List<SyncBitmapInfo>> CreateImage(ImageFormat imageFormat, IPageFilter pageFilter,
            int resolution)
        {
            var list = new List<byte[]>();
            var syncBitmapInfos = new List<SyncBitmapInfo>();
            foreach (var sectionGroupFunc in SectionGroups)
            {
                var sectionGroup = sectionGroupFunc();
                if (sectionGroup.Sections.Count == 0) continue;
                var pages = new LinkedList<byte[]>();
                var pageMustBeAdd = pageFilter.PageMustBeAdd;
                var pageTuples = FillBitmap(xGraphics => TableRenderer.Draw(xGraphics,
                        (pageIndex, action, section) => {
                            FillBitmap(graphics => {
                                    action(graphics);
                                    return new { };
                                },
                                bitmap => {
                                    if (pageFilter.PageMustBeAdd)
                                        pages.AddLast(ToBytes(bitmap, imageFormat));
                                },
                                section.PageSettings, resolution);
                        }, this, GraphicsType.Image, sectionGroup, pageFilter),
                    bitmap => {
                        if (pageMustBeAdd)
                            pages.AddFirst(ToBytes(bitmap, imageFormat));
                    },
                    sectionGroup.Sections[0].PageSettings, resolution);
                syncBitmapInfos.AddRange(pageTuples.Select(pageTuple => new SyncBitmapInfo {
                    PageInfo = pageTuple.SyncPageInfo,
                    Resolution = resolution,
                    HorizontalPixelCount = HorizontalPixelCount(pageTuple.Section.PageSettings, resolution),
                    VerticalPixelCount = VerticalPixelCount(pageTuple.Section.PageSettings, resolution),
                }));
                list.AddRange(pages);
            }
            return Tuple.Create(list, syncBitmapInfos);
        }

        public string SavePng(int pageNumber, string path, int resolution = defaultResolution)
        {
            var tuple = CreatePng(resolution, new OnePageFilter(pageNumber));
            File.WriteAllBytes(path, tuple.Item1[0]);
            File.WriteAllText(Path.ChangeExtension(path, ".json"),
                JsonConvert.SerializeObject(tuple.Item2[0], Formatting.Indented));
            return path;
        }
        
        private const int defaultResolution = 254;

        public static T FillBitmap<T>(Func<IGraphics, T> func, Action<Bitmap> action2, PageSettings pageSettings, int resolution)
        {
            var horizontalPixelCount = HorizontalPixelCount(pageSettings, resolution);
            var verticalPixelCount = VerticalPixelCount(pageSettings, resolution);
            using (var bitmap = new Bitmap(horizontalPixelCount, verticalPixelCount))
            {
                T result;
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, horizontalPixelCount, verticalPixelCount);
                    var s = (float) (resolution / 72d);
                    graphics.ScaleTransform(s, s);
                    result = func(new ImageGraphics(graphics));
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