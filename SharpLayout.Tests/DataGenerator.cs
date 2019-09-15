using System;
using System.IO;
using Newtonsoft.Json;
using Xunit;

namespace SharpLayout.Tests
{
    public class DataGenerator
    {
        [Fact]
        public void PaymentOrder()
        {
            WriteData(new PaymentData {IncomingDate = DateTime.Now, OutcomingDate = DateTime.Now});
        }

        private static void WriteData(PaymentData data)
        {
            File.WriteAllText($"{data.GetType().FullName}.json", 
                JsonConvert.SerializeObject(data, Formatting.Indented));
        }
    }
}