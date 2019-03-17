using System;

namespace SharpLayout.Tests
{
    static class Program
    {
        static void Main()
        {
            var document = new Document();
            PaymentOrder.AddSection(document, new PaymentData {
                IncomingDate = DateTime.Now,
                OutcomingDate = DateTime.Now
            });
            document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf");
        }
    }
}
