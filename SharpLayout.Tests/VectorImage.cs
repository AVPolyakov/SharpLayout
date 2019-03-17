using System;
using System.IO;
using PdfSharp.Drawing;

namespace SharpLayout.Tests
{
    public static class VectorImage
    {
        public static void AddSection(Document document)
        {
            var section = document.Add(new Section(new PageSettings()));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Content(new VectorImageContent()));
            }
        }

        public class VectorImageContent : IImageContent
        {
            public static readonly Lazy<ImageInfo> ImageInfo = new Lazy<ImageInfo>(() => {
                byte[] bytes;
                var anchorType = typeof(Images.Images);
                using (var stream = anchorType.Assembly
                    .GetManifestResourceStream($"{anchorType.Namespace}.blue-rabbit.pdf"))
                using (var memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    bytes = memoryStream.ToArray();
                }
                using (var stream = new MemoryStream(bytes))
                using (var xImage = XImage.FromStream(stream))
                    return new ImageInfo(bytes, xImage.PointWidth, xImage.PointHeight);
            });

            public Stream CreateStream() => new MemoryStream(ImageInfo.Value.Bytes);
        }

    }

    public class ImageInfo
    {
        public byte[] Bytes { get; }
        public double PointWidth { get; }
        public double PointHeight { get; }

        public ImageInfo(byte[] bytes, double pointWidth, double pointHeight)
        {
            Bytes = bytes;
            PointWidth = pointWidth;
            PointHeight = pointHeight;
        }
    }
}