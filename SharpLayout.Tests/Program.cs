using System;
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
            public const string TimesNewRoman = "TimesNewRoman";
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            switch (familyName)
            {
                case FamilyNames.TimesNewRoman:
                    return new FontResolverInfo(FaceNames.TimesNewRoman);
            }
            throw new Exception("Font not found");
        }

        public byte[] GetFont(string faceName)
        {
            throw new NotImplementedException();
        }
    }
}
