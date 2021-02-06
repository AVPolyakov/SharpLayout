using System.Collections.Concurrent;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Font
    {
        private static readonly ConcurrentDictionary<FontKey, XFont> cache = new();

        private static XFont GetXFont(FontFamilyInfo familyInfo, double emSize, XFontStyle style, XPdfFontOptions pdfOptions, out  FontKey key)
        {
            key = new FontKey(familyInfo.FullName, emSize, style, pdfOptions.FontEncoding);
            if (cache.TryGetValue(key, out var value))
                return value;
            var xFont = new XFont(familyInfo.FullName, emSize, style, pdfOptions);
            cache.TryAdd(key, xFont);
            return xFont;
        }

        public FontKey Key { get; }

        public Font(FontFamilyInfo familyInfo, double emSize, XFontStyle style, XPdfFontOptions pdfOptions)
        {
            FamilyInfo = familyInfo;
            XFont = GetXFont(familyInfo, emSize, style, pdfOptions, out  var key);
            Key = key;
        }

        public FontFamilyInfo FamilyInfo { get; }
        
        public XFont XFont { get; }

        public XFontFamily FontFamily => XFont.FontFamily;

        public double GetHeight() => XFont.GetHeight();

        public bool Underline => XFont.Underline;

        public double Size => XFont.Size;

        public XFontStyle Style => XFont.Style;

        public XPdfFontOptions PdfOptions => XFont.PdfOptions;

        public bool Bold => XFont.Bold;

        public string Name => XFont.Name;
    }
}