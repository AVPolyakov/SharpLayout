using System;
using System.Linq;
using System.Threading.Tasks;

namespace SharpLayout.Tests
{
    static class Program
    {
        static async Task Main()
        {
            for (var k = 0; k < 1000; k++)
            {
                await Task.WhenAll(Enumerable.Range(1, 16).Select(j => Task.Run(() => {
                    var document = new Document();
                    for (var i = 0; i < 20; i++)
                    {
                        PaymentOrder.AddSection(document,
                            new PaymentData {IncomingDate = DateTime.Now, OutcomingDate = DateTime.Now});
                    }
                    document.CreatePdf();
                })));
                Console.WriteLine(k);
            }
        }
    }
}
