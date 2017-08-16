using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout.Tests
{
    public static class Styles
    {
        public static Paragraph TimesNewRoman10_5(string text) =>
            new Paragraph().Add(new Span(text, new XFont("Times New Roman", 10.5, XFontStyle.Regular, PdfOptions)));

        public static Paragraph TimesNewRoman9_5(string text) =>
            new Paragraph().Add(new Span(text, new XFont("Times New Roman", 9.5, XFontStyle.Regular, PdfOptions)));

        public static Paragraph TimesNewRoman11_5Bold(string text) =>
            new Paragraph().Add(new Span(text, new XFont("Times New Roman", 11.5, XFontStyle.Bold, PdfOptions)));

        public static Paragraph TimesNewRoman9(string text) =>
            new Paragraph().Add(new Span(text, new XFont("Times New Roman", 9, XFontStyle.Regular, PdfOptions)));

        public const double BorderWidth = 0.5D;

        public static XPdfFontOptions PdfOptions => new XPdfFontOptions(PdfFontEncoding.Unicode);
    }
}