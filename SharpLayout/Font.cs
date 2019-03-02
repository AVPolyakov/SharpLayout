using System;
using System.Collections.Concurrent;
using System.Drawing;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Font
    {
        private static readonly ConcurrentDictionary<FontKey, XFont> cache =
            new ConcurrentDictionary<FontKey, XFont>();

        private static XFont GetXFont(string familyName, double emSize, XFontStyle style, XPdfFontOptions pdfOptions)
        {
            var key = new FontKey(familyName, emSize, style, pdfOptions.FontEncoding);
            if (cache.TryGetValue(key, out var value))
                return value;
            var xFont = new XFont(familyName, emSize, style, pdfOptions);
            cache.TryAdd(key, xFont);
            return xFont;
        }

        public Font(string familyName, double emSize, XFontStyle style, XPdfFontOptions pdfOptions)
        {
            gdiFont = Lazy.Create(() => new System.Drawing.Font(
                Name, (float)Size, (FontStyle)Style, GraphicsUnit.World));
            XFont = GetXFont(familyName, emSize, style, pdfOptions);
        }

        public XFont XFont { get; }

        public XFontFamily FontFamily => XFont.FontFamily;

        public double GetHeight() => XFont.GetHeight();

        public bool Underline => XFont.Underline;

        public double Size => XFont.Size;

        public XFontStyle Style => XFont.Style;

        public XPdfFontOptions PdfOptions => XFont.PdfOptions;

        public bool Bold => XFont.Bold;

        public string Name => XFont.Name;

        private readonly Lazy<System.Drawing.Font> gdiFont;

        public System.Drawing.Font GdiFont => gdiFont.Value;
    }
}