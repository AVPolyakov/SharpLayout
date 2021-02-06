using System;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using SharpLayout;
using static Resources.FontFamilies;
using static SharpLayout.Direction;
using static SharpLayout.Util;

namespace Examples
{
    public static class Styles
    {
        public static Font TimesNewRoman10_5 => new Font(TimesNewRoman, 10.5, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman60Bold => new Font(TimesNewRoman, 60, XFontStyle.Bold, PdfOptions);

        public static Font TimesNewRoman7_5 => new Font(TimesNewRoman, 7.5, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman8 => new Font(TimesNewRoman, 8, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman9 => new Font(TimesNewRoman, 9, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman9_5 => new Font(TimesNewRoman, 9.5, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman10 => new Font(TimesNewRoman, 10, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman10Bold => new Font(TimesNewRoman, 10, XFontStyle.Bold, PdfOptions);
        public static Font TimesNewRoman10_5Bold => new Font(TimesNewRoman, 10.5, XFontStyle.Bold, PdfOptions);
        public static Font TimesNewRoman11 => new Font(TimesNewRoman, 11, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman11_5 => new Font(TimesNewRoman, 11.5, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman11_5Bold => new Font(TimesNewRoman, 11.5, XFontStyle.Bold, PdfOptions);
        public static Font TimesNewRoman12 => new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman14 => new Font(TimesNewRoman, 14, XFontStyle.Regular, PdfOptions);
        public static Font Courier15 => new Font(CourierNew, 15, XFontStyle.Regular, PdfOptions);
        public static Font TimesNewRoman12Bold => new Font(TimesNewRoman, 12, XFontStyle.Bold, PdfOptions);
        public static Font TimesNewRoman13Bold => new Font(TimesNewRoman, 13, XFontStyle.Bold, PdfOptions);
        
        public static Font Arial9 => new Font(Arial, 9, XFontStyle.Regular, PdfOptions);

        public const double BorderWidth = 0.5D;

        public static Cell Border(this Cell cell, Direction direction) => cell.Border(direction, BorderWidth);

        public static XPdfFontOptions PdfOptions => new XPdfFontOptions(PdfFontEncoding.Unicode);

        public static string Format(this DateTime value) => value.ToString("dd.MM.yyyy");
        
        public static void Distribute(this Column[] columns, double width)
        {
            var d = width / columns.Length;
            double sum = 0;
            for (var i = 0; i < columns.Length - 1; i++)
            {
                columns[i].Width += d;
                sum += d;
            }
            columns[columns.Length - 1].Width += width - sum;
        }
        
        public static string CellSubstring(this string s, int i, int cellCount)
        {
            if (s == null) return "";
            if (i >= s.Length) return "";
            if (i == cellCount - 1) return s.Substring(i);
            return s.Substring(i, 1);
        }
        
        public static Paragraph NormalParagraph => new Paragraph().Margin(Left | Right, Cm(0.05));
    }
}