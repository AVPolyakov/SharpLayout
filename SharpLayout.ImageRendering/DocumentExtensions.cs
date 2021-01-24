using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using PdfSharp.Drawing;

namespace SharpLayout.ImageRendering
{
    public static class DocumentExtensions
    {
        public static Tuple<List<byte[]>, List<SyncBitmapInfo>> CreatePng(this Document document, double resolution = defaultResolution)
        {
            return document.CreatePng(resolution, AllPageFilter.Instance);
        }
        
        private static Tuple<List<byte[]>, List<SyncBitmapInfo>> CreatePng(this Document document, double resolution, IPageFilter pageFilter)
        {
            return document.CreateImage(ImageFormat.Png, pageFilter, resolution);
        }

        public static Tuple<List<byte[]>, List<SyncBitmapInfo>> CreateImage(this Document document, ImageFormat imageFormat, double resolution = defaultResolution)
        {
            return document.CreateImage(imageFormat, AllPageFilter.Instance, resolution);
        }

        private static Tuple<List<byte[]>, List<SyncBitmapInfo>> CreateImage(this Document document, ImageFormat imageFormat, IPageFilter pageFilter,
            double resolution)
        {
            var list = new List<byte[]>();
            var syncBitmapInfos = new List<SyncBitmapInfo>();
            foreach (var sectionGroupFunc in document.SectionGroups)
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
                                section, resolution);
                        }, document, GraphicsType.Image, sectionGroup, pageFilter),
                    bitmap => {
                        if (pageMustBeAdd)
                            pages.AddFirst(ToBytes(bitmap, imageFormat));
                    },
                    sectionGroup.Sections[0], resolution);
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

        public static string SavePng(this Document document, int pageNumber, string path, double resolution = defaultResolution)
        {
            var tuple = document.CreatePng(resolution, new OnePageFilter(pageNumber));
            File.WriteAllBytes(path, tuple.Item1[0]);
            File.WriteAllText(Path.ChangeExtension(path, ".json"),
                JsonConvert.SerializeObject(tuple.Item2[0], Formatting.Indented));
            return path;
        }
        
        private const double defaultResolution = 254;

        private static T FillBitmap<T>(Func<IGraphics, T> func, Action<Bitmap> action2, Section section, double resolution)
        {
            var horizontalPixelCount = HorizontalPixelCount(section.PageSettings, resolution);
            var verticalPixelCount = VerticalPixelCount(section.PageSettings, resolution);
            Bitmap bitmap;
            if (section.Template().HasValue)
                using (var image = System.Drawing.Image.FromStream(section.Template().Value()))
                    bitmap = new Bitmap(image, horizontalPixelCount, verticalPixelCount);
            else
                bitmap = new Bitmap(horizontalPixelCount, verticalPixelCount);
            using (bitmap)
            {
                T result;
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    if (!section.Template().HasValue)
                        graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, horizontalPixelCount, verticalPixelCount);
                    var s = (float) (resolution / 72d);
                    graphics.ScaleTransform(s, s);
                    result = func(new ImageGraphics(graphics));
                }
                bitmap.SetResolution((float) resolution, (float) resolution);
                action2(bitmap);
                return result;
            }
        }
        
        private static int VerticalPixelCount(PageSettings pageSettings, double resolution) 
            => (int) (new XUnit(pageSettings.PageHeight).Inch * resolution);

        private static int HorizontalPixelCount(PageSettings pageSettings, double resolution) 
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