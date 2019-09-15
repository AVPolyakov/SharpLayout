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

        [Fact]
        public void ContractDealPassport()
        {
            WriteData(new ContractDealPassportData());
        }

        [Fact]
        public void LoanAgreementDealPassport()
        {
            WriteData(new LoanAgreementDealPassportData());
        }

        private static void WriteData(object data)
        {
            File.WriteAllText($"{data.GetType().FullName}.json", 
                JsonConvert.SerializeObject(data, Formatting.Indented));
        }
    }
}