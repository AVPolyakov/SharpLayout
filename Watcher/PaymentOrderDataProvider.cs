using System;
using Examples;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class PaymentOrderDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            return new PaymentOrderQuery().Get(1);
        }
    }
}