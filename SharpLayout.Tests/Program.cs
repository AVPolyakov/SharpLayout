using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using PdfSharp.Fonts;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            GlobalFontSettings.FontResolver = new FontResolver();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var document = new Document();
            PaymentOrder.AddSection(document, new PaymentData {
                IncomingDate = DateTime.Now,
                OutcomingDate = DateTime.Now,
                PaymentPurpose = RuntimeInformation.OSDescription
            });
            document.SavePdf("Temp.pdf");
        }
    }

    public class FontResolver : IFontResolver
    {
        private static class FamilyNames
        {
            public const string TimesNewRoman = "Times New Roman";
        }

        private static class FaceNames
        {            
            public const string TimesNewRomanBoldItalic = "TimesNewRomanBoldItalic";
            public const string TimesNewRomanBold = "TimesNewRomanBold";
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            switch (familyName)
            {
                case FamilyNames.TimesNewRoman:
                    if (isBold)
                    {
                        if (isItalic)
                            return new FontResolverInfo(FaceNames.TimesNewRomanBoldItalic);
                        else
                            return new FontResolverInfo(FaceNames.TimesNewRomanBold);
                    }
                    
            }
            throw new Exception("Font not found");
        }

        public byte[] GetFont(string faceName)
        {
            string fileName;
            switch (faceName)
            {
                case FaceNames.TimesNewRomanBoldItalic:
                    fileName = @"timesbi.ttf";
                    break;
                case FaceNames.TimesNewRomanBold:
                    fileName = @"timesbd.ttf";
                    break;
                default:
                    throw new Exception("Font file not found");
            }
            return File.ReadAllBytes($@"Fonts\{fileName}");
        }
    }
}
