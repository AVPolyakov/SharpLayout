using System;
using System.Runtime.InteropServices;
using System.Text;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
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
}
