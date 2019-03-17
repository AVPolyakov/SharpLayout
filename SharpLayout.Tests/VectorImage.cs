using System;
using System.IO;
using PdfSharp.Drawing;
using SharpLayout.Tests.Images;
using static SharpLayout.Direction;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class VectorImage
    {
        public static void AddSection(Document document)
        {
            var section = document.Add(new Section(new PageSettings()));
            var info = VectorImageContent.ImageInfo.Value;
            {
                var table = section.AddTable();
                var scale = 0.7;
                var c1 = table.AddColumn(info.PointWidth * scale);
                table.AddColumn(Px(500));
                table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(new VectorImageContent()));
            }
            {
                var table = section.AddTable().Margin(Top, Cm(0.5));
                var scale = 0.3;
                var c1 = table.AddColumn();
                var c2 = table.AddColumn();
                var c3 = table.AddColumn();
                table.Columns.ToArray().Distribute(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(new VectorImageContent()));
                var r2 = table.AddRow();
                r2[c2].Add(new Image().Alignment(HorizontalAlign.Center)
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(new VectorImageContent()));
                var r3 = table.AddRow();
                r3[c3].Add(new Image().Alignment(HorizontalAlign.Right)
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(new VectorImageContent()));
            }
        }

        public class VectorImageContent : IImageContent
        {
            public static readonly Lazy<ImageInfo> ImageInfo = new Lazy<ImageInfo>(() => {
                byte[] bytes;
                var anchorType = typeof(BlueRabbit);
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