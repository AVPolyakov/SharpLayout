using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Drawing2D;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace SharpLayout
{
    public interface IGraphics
    {
        void DrawImage(IImageContent imageContent, double x, double y);
        void DrawImage(IImageContent imageContent, double x, double y, Func<double, double> width, Func<double, double> height);
        void DrawString(string s, Font font, XBrush brush, double x, double y);
        void DrawLine(XPen pen, double x1, double y1, double x2, double y2);
        void DrawRectangle(XBrush brush, double x, double y, double width, double height);
        XSize MeasureString(string text, XFont font, XStringFormat stringFormat);
    }

    public class ImageGraphics : IGraphics
    {
        private readonly System.Drawing.Graphics graphics;

        public ImageGraphics(System.Drawing.Graphics graphics)
        {
            this.graphics = graphics;
        }

        public void DrawImage(IImageContent imageContent, double x, double y)
        {
            using (var stream = imageContent.CreateStream())
                if (PdfReader.TestPdfFile(stream) <= 0)
                    using (var image = System.Drawing.Image.FromStream(stream))
                        graphics.DrawImage(image, (float) x, (float) y,
                            GetDistance(image, image.Width),
                            GetDistance(image, image.Height));
                else
                    using (var image = XImage.FromStream(stream))
                        DrawMissingImage(x, y,
                            image.PointWidth,
                            image.PointHeight);
        }

        public void DrawImage(IImageContent imageContent, double x, double y,
            Func<double, double> width, Func<double, double> height)
        {
            using (var stream = imageContent.CreateStream())
                if (PdfReader.TestPdfFile(stream) <= 0)
                    using (var image = System.Drawing.Image.FromStream(stream))
                        graphics.DrawImage(image, (float) x, (float) y,
                            (float) width(GetDistance(image, image.Width)),
                            (float) height(GetDistance(image, image.Height)));
                else
                    using (var image = XImage.FromStream(stream))
                        DrawMissingImage(x, y,
                            width(image.PointWidth),
                            height(image.PointHeight));
        }

        private static float GetDistance(System.Drawing.Image image, int distance)
        {
            return distance * 72 / image.HorizontalResolution;
        }

        public void DrawString(string s, Font font, XBrush brush, double x, double y)
        {
            var lineSpace = font.XFont.GetHeight();
            var cyAscent = lineSpace * font.XFont.CellAscent / font.XFont.CellSpace;
            graphics.DrawString(s,
                font.GdiFont,
                GetBrush(brush),
                (float) x,
                (float) (y - cyAscent),
                drawStringFormat);
        }

        private static readonly StringFormat drawStringFormat = GetStringFormat();

        private static StringFormat GetStringFormat()
        {
            var getStringFormat = (StringFormat) StringFormat.GenericTypographic.Clone();
            // Bugfix: Set MeasureTrailingSpaces to get the correct width with Graphics.MeasureString().
            // Before, MeasureString() didn't include blanks in width calculation, which could result in text overflowing table or page border before wrapping. $MaOs
            getStringFormat.FormatFlags = getStringFormat.FormatFlags | StringFormatFlags.MeasureTrailingSpaces;
            return getStringFormat;
        }

        private static SolidBrush GetBrush(XBrush brush)
        {
            if (brush is XSolidBrush solidBrush)
            {
                var color = solidBrush.Color;
                var key = new ColorKey(color.A, color.R, color.G, color.B);
                if (brushCache.TryGetValue(key, out var value))
                    return value;
                var result = new SolidBrush(ToGdiColor(color));
                brushCache.TryAdd(key, result);
                return result;
            }
            throw new InvalidOperationException("Implemented for XSolidBrush only.");
        }

        private static Color ToGdiColor(XColor color)
        {
            return Color.FromArgb((int) (color.A * 255), color.R, color.G, color.B);
        }

        public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
        {
            graphics.DrawLine(GetPen(pen),
                (float) x1, (float) y1, (float) x2, (float) y2);
        }

        public static LineCap ToLineCap(XLineCap lineCap)
        {
            return gdiLineCap[(int)lineCap];
        }

        private static readonly LineCap[] gdiLineCap = {LineCap.Flat, LineCap.Round, LineCap.Square};

        public static LineJoin ToLineJoin(XLineJoin lineJoin)
        {
            return gdiLineJoin[(int)lineJoin];
        }

        private static readonly LineJoin[] gdiLineJoin = {LineJoin.Miter, LineJoin.Round, LineJoin.Bevel};

        private static Pen GetPen(XPen xPen)
        {
            var pen = new Pen(ToGdiColor(xPen.Color), (float) xPen.Width);
            var lineCap = ToLineCap(xPen.LineCap);
            pen.StartCap = lineCap;
            pen.EndCap = lineCap;
            pen.LineJoin = ToLineJoin(xPen.LineJoin);
            pen.DashOffset = (float) xPen.DashOffset;
            if (xPen.DashStyle == XDashStyle.Custom)
            {
                var len = xPen.DashPattern == null ? 0 : xPen.DashPattern.Length;
                var pattern = new float[len];
                for (var idx = 0; idx < len; idx++)
                    pattern[idx] = (float) xPen.DashPattern[idx];
                pen.DashPattern = pattern;
            }
            else
                pen.DashStyle = (DashStyle) xPen.DashStyle;
            return pen;
        }

        public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
        {
            graphics.FillRectangle(GetBrush(brush), (float) x, (float) y, (float) width, (float) height);
        }

        public XSize MeasureString(string text, XFont font, XStringFormat stringFormat)
        {
            return pdfGraphics.MeasureString(text, font, stringFormat);
        }

        private static readonly Graphics pdfGraphics = GetPdfGraphics();

        private static Graphics GetPdfGraphics()
        {
            return new Graphics(XGraphics.FromPdfPage(new PdfDocument().AddPage()));
        }

        private void DrawMissingImage(double x, double y, double width, double height)
        {
            var fX = (float) x;
            var fY = (float) y;
            var fWidth = (float) width;
            var fHeight = (float) height;
            var pen = Pens.Red;
            graphics.DrawRectangle(pen, fX, fY, fWidth, fHeight);
            graphics.DrawLine(pen, fX, fY, fX + fWidth, fY + fHeight);
            graphics.DrawLine(pen, fX + fWidth, fY, fX, fY + fHeight);
        }

        private static readonly ConcurrentDictionary<ColorKey, SolidBrush> brushCache =
            new ConcurrentDictionary<ColorKey, SolidBrush>();

        private struct ColorKey : IEquatable<ColorKey>
        {
            private double A { get; }
            private byte R { get; }
            private byte G { get; }
            private byte B { get; }

            public ColorKey(double a, byte r, byte g, byte b)
            {
                A = a;
                R = r;
                G = g;
                B = b;
            }

            public bool Equals(ColorKey other)
            {
                return A.Equals(other.A) && R == other.R && G == other.G && B == other.B;
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                return obj is ColorKey other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    var hashCode = A.GetHashCode();
                    hashCode = (hashCode * 397) ^ R.GetHashCode();
                    hashCode = (hashCode * 397) ^ G.GetHashCode();
                    hashCode = (hashCode * 397) ^ B.GetHashCode();
                    return hashCode;
                }
            }
        }
    }

    public class Graphics: IGraphics
    {
        private readonly XGraphics xGraphics;

        public Graphics(XGraphics xGraphics)
        {
            this.xGraphics = xGraphics;
        }

        public void DrawImage(IImageContent imageContent, double x, double y)
        {
            using (var stream = imageContent.CreateStream())
            using (var xImage = XImage.FromStream(stream))
                xGraphics.DrawImage(xImage, x, y);
        }

        public void DrawImage(IImageContent imageContent, double x, double y, 
            Func<double, double> width, Func<double, double> height)
        {
            using (var stream = imageContent.CreateStream())
            using (var xImage = XImage.FromStream(stream))
                xGraphics.DrawImage(xImage, x, y, width(xImage.PointWidth), height(xImage.PointHeight));
        }

        public void DrawString(string s, Font font, XBrush brush, double x, double y)
        {
            xGraphics.DrawString(s, font.XFont, brush, x, y);
        }

        public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
        {
            xGraphics.DrawLine(pen, x1, y1, x2, y2);
        }

        public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
        {
            xGraphics.DrawRectangle(brush, x, y, width, height);
        }

        public XSize MeasureString(string text, XFont font, XStringFormat stringFormat)
        {
            return xGraphics.MeasureString(text, font, stringFormat);
        }
    }
}