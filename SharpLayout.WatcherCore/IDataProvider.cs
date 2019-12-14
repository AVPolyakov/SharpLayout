using System;

namespace SharpLayout.WatcherCore
{
    public interface IDataProvider
    {
        object Create(Func<object> deserializeFunc);
    }
}