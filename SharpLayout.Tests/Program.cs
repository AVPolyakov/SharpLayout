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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            GlobalFontSettings.FontResolver = new FontResolver();

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
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            throw new NotImplementedException();
        }

        public byte[] GetFont(string faceName)
        {
            throw new NotImplementedException();
        }
    }
}
