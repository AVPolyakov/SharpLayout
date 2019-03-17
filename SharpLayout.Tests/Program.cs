using System;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var document = new Document();
            PaymentOrder.AddSection(document, new PaymentData {
                IncomingDate = DateTime.Now,
                OutcomingDate = DateTime.Now,
                PaymentPurpose = @"Тест
В том числе НДС 270,00"
            });
            document.SavePdf("Temp.pdf");
        }
    }
}
