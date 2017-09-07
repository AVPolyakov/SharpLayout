using System;
using System.Collections.Generic;

namespace SharpLayout
{
    public class Lazy<TKey, TValue>
    {
        private readonly Dictionary<TKey, TValue> dictionary;
        private readonly Func<TKey, TValue> func;

        public Lazy(Dictionary<TKey, TValue> dictionary, Func<TKey, TValue> func)
        {
            this.dictionary = dictionary;
            this.func = func;
        }

        public TValue GetValue(TKey key)
        {
            {
                TValue value;
                if (dictionary.TryGetValue(key, out value)) return value;
            }
            {
                var value = func(key);
                dictionary.Add(key, value);
                return value;
            }
        }
    }

    public static class Lazy
    {
        public static Lazy<TKey, TValue> Create<TKey, TValue>(Dictionary<TKey, TValue> dictionary, Func<TKey, TValue> func)
            => new Lazy<TKey, TValue>(dictionary, func);
    }
}