using System;
using Examples;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class PaymentOrderDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            var data = (PaymentOrderData) deserializeFunc();
            
            data.IncomingDate = new DateTime(2019, 12, 14);
            data.OutcomingDate = new DateTime(2019, 12, 17);
            
            return data;
        }
    }
}