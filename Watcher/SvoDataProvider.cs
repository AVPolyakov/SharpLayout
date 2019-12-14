using System;
using SharpLayout.WatcherCore;

namespace Watcher
{
    public class SvoDataProvider : IDataProvider
    {
        public object Create(Func<object> deserializeFunc)
        {
            return deserializeFunc();
        }
    }
}