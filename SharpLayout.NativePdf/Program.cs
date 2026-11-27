using System;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using SharpLayout;
using SharpLayout.NativePdf;
using static Resources.FontFamilies;

namespace MeasureComparisonTest
{
    internal class Program
    {
        private static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            M2();
        }

        private static void M2()
        {
            var doc = new MinimalPdfDocument();
            var font = doc.GetOrRegisterFont(new Font(TimesNewRoman, 8, XFontStyle.Regular, PdfOptions));
            var page = doc.AddPage();

            page.DrawString("Привет, мир! Платёжное поручение № 104 от 18.11.2026", font, 12, Colors.Black, 50, 50);

            using (var fs = File.Create("only_text.pdf"))
            {
                doc.Save(fs);
            }

            Console.WriteLine("Файл only_text.pdf успешно создан.");
        }
        
        public static XPdfFontOptions PdfOptions => new(PdfFontEncoding.Unicode);
        
        private static void M1()
        {
            byte[] ttfBytes = File.ReadAllBytes("Roboto-Regular.ttf");
            string sample = "Платёжное поручение № 104 от 17.11.2026";
            double fontSize = 10.0;

            // 1. Замер через наш новый класс TrueTypeMetrics
            var metrics = new TrueTypeMetrics(ttfBytes, false);
            LayoutSize nativeSize = metrics.MeasureString(sample, fontSize);

            // 2. Замер через ваш текущий форк PDFsharp
            GlobalFontSettings.FontResolver = new SingleFontResolver(ttfBytes);
            var pdfDoc = new PdfDocument();
            var page = pdfDoc.AddPage();
            using (var xGraphics = XGraphics.FromPdfPage(page))
            {
                var xFont = new XFont("Roboto", fontSize, XFontStyle.Regular);
                XSize pdfSharpSize = xGraphics.MeasureString(sample, xFont);

                // 3. Вывод результатов сравнения
                Console.WriteLine($"Строка: \"{sample}\" (кегль {fontSize} pt)");
                Console.WriteLine($"GID 'A': {metrics.GetGlyphId('A')}, GID 'П': {metrics.GetGlyphId('П')}");
                Console.WriteLine(new string('-', 55));
                Console.WriteLine($"PDFsharp:        Width = {pdfSharpSize.Width:F6} pt | Height = {pdfSharpSize.Height:F6} pt");
                Console.WriteLine($"TrueTypeMetrics: Width = {nativeSize.Width:F6} pt | Height = {nativeSize.Height:F6} pt");
                Console.WriteLine(new string('-', 55));
                Console.WriteLine($"Разница:         dWidth = {Math.Abs(pdfSharpSize.Width - nativeSize.Width):F6} | dHeight = {Math.Abs(pdfSharpSize.Height - nativeSize.Height):F6}");
            }
        }
    }

    // Простой FontResolver для загрузки Roboto-Regular.ttf из памяти в PDFsharp
    internal class SingleFontResolver : IFontResolver
    {
        private readonly byte[] _fontData;
        public SingleFontResolver(byte[] fontData) => _fontData = fontData;

        public string DefaultFontName => "Roboto";

        public byte[] GetFont(string faceName) => _fontData;

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            return new FontResolverInfo("Roboto#Regular");
        }
    }
}
