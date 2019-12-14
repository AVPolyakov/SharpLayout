using System;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class LoanAgreementDealPassportDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            return deserializeFunc();
        }
    }
}