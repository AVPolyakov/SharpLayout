using System;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class ContractDealPassportDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            return deserializeFunc();
        }
    }
}