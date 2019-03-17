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
            Svo.AddSection(document);
            ContractDealPassport.AddSection(document);
            LoanAgreementDealPassport.AddSection(document);
            VectorImage.AddSection(document);
            PngImage.AddSection(document);
            document.SavePdf("Temp.pdf");
        }
    }
}
