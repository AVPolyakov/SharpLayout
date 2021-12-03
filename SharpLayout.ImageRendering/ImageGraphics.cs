using System;
using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace SharpLayout.ImageRendering
{
    public class ImageGraphics : IGraphics
    {
        private readonly System.Drawing.Graphics graphics;

        public ImageGraphics(System.Drawing.Graphics graphics)
        {
            this.graphics = graphics;
        }

        public void DrawImage(Func<Stream> imageContent, double x, double y)
        {
            using (var stream = imageContent())
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

        public void DrawImage(Func<Stream> imageContent, double x, double y,
            Func<double, double> width, Func<double, double> height)
        {
            using (var stream = imageContent())
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
                GetGdiFont(font),
                GetBrush(brush),
                (float) x,
                (float) (y - cyAscent),
                drawStringFormat);
        }

        private static readonly ConcurrentDictionary<FontResolutionKey, FontFamily> fontFamilies = new ();

        private static System.Drawing.Font GetGdiFont(Font font)
        {
            var key = font.Key;
            if (gdiFontCache.TryGetValue(key, out var value))
                return value;

            var fontResolutionKey = new FontResolutionKey(font.FamilyInfo.Name,
                isBold: font.Style.HasFlag(XFontStyle.Bold), isItalic: font.Style.HasFlag(XFontStyle.Italic));
            if (!fontFamilies.ContainsKey(fontResolutionKey))
            {
                var fontResolverInfo = GlobalFontSettings.FontResolver.ResolveTypeface(font.FamilyInfo.Name,
                    isBold: font.Style.HasFlag(XFontStyle.Bold), isItalic: font.Style.HasFlag(XFontStyle.Italic));
                var bytes = GlobalFontSettings.FontResolver.GetFont(fontResolverInfo.FaceName);

                var fontCollection = new PrivateFontCollection();

                var fontDataPtr = Marshal.AllocCoTaskMem(bytes.Length);
                Marshal.Copy(bytes, 0, fontDataPtr, bytes.Length);
                fontCollection.AddMemoryFont(fontDataPtr, bytes.Length);

                var family = fontCollection.Families[0];

                fontFamilies.TryAdd(fontResolutionKey, family);
            }

            var gdiFont = new System.Drawing.Font(
                fontFamilies[fontResolutionKey], (float) font.Size, (FontStyle) font.Style, GraphicsUnit.World);
            gdiFontCache.TryAdd(key, gdiFont);
            return gdiFont;
        }
        
        private static readonly ConcurrentDictionary<FontKey, System.Drawing.Font> gdiFontCache =
            new ConcurrentDictionary<FontKey, System.Drawing.Font>();

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

        public void DrawLines(XPen pen, XPoint[] points)
        {
            graphics.DrawLines(GetPen(pen), points.Select(_ => new PointF((float)_.X, (float)_.Y)).ToArray());
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
}