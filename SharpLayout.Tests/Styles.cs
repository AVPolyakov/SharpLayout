using System;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout.Tests
{
    public static class Styles
    {
        public static XFont TimesNewRoman10_5 => new XFont("Times New Roman", 10.5, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman60Bold => new XFont("Times New Roman", 60, XFontStyle.Bold, PdfOptions);

        public static XFont TimesNewRoman8 => new XFont("Times New Roman", 8, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman9 => new XFont("Times New Roman", 9, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman9_5 => new XFont("Times New Roman", 9.5, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman10 => new XFont("Times New Roman", 10, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman10Bold => new XFont("Times New Roman", 10, XFontStyle.Bold, PdfOptions);
        public static XFont TimesNewRoman10_5Bold => new XFont("Times New Roman", 10.5, XFontStyle.Bold, PdfOptions);
        public static XFont TimesNewRoman11 => new XFont("Times New Roman", 11, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman11_5 => new XFont("Times New Roman", 11.5, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman11_5Bold => new XFont("Times New Roman", 11.5, XFontStyle.Bold, PdfOptions);
        public static XFont TimesNewRoman12 => new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions);
        public static XFont TimesNewRoman12Bold => new XFont("Times New Roman", 12, XFontStyle.Bold, PdfOptions);
        public static XFont TimesNewRoman13Bold => new XFont("Times New Roman", 13, XFontStyle.Bold, PdfOptions);

        public const double BorderWidth = 0.5D;

        public static Cell Border(this Cell cell, Direction direction) => cell.Border(direction, BorderWidth);

        public static XPdfFontOptions PdfOptions => new XPdfFontOptions(PdfFontEncoding.Unicode);

        public static string FormatDate(DateTime value) => value.ToString("dd.MM.yyyy");
    }
}