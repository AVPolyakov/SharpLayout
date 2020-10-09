using System;
using Examples;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class LegalEntityCreationDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            return new LegalEntityCreationData
            {
                FullName = "Полное имя aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                Name = "Сокращенное имя",
                PostalCode = "123456",
                Subject = "АБ",
                AreaName = "Большой улус аааааааааааааааааааааааааааааааа",
                Area = "Улус №1",
                City = "Город №1",
                CityName = "Большой город",
            };
        }
    }
}