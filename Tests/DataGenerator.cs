using System;
using System.IO;
using Examples;
using Newtonsoft.Json;
using Xunit;

namespace Tests
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

        [Fact]
        public void Svo()
        {
            WriteData(new SvoData());
        }

        [Fact]
        public void LegalEntityCreationData()
        {
            WriteData(new LegalEntityCreationData
            {
                FullName = "Полное имя aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                Name = "Сокращенное имя",
                PostalCode = "123456",
                Subject = "АБ",
                AreaName = "Большой улус аааааааааааааааааааааааааааааааа",
                Area = "Улус №1",
                City = "Город №1",
                CityName = "Большой город",
            });
        }

        private static void WriteData(object data)
        {
            File.WriteAllText($@"..\..\..\..\Starter\bin\Debug\netcoreapp3.1\{data.GetType().FullName}_DataSource.json", 
                JsonConvert.SerializeObject(data, Formatting.Indented));
        }
    }
}