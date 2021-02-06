using System;

namespace SharpLayout
{
    public interface IHandler<out T>
    {
        TResult Handle<TResult>(Func<T, TResult> func);
    }
}