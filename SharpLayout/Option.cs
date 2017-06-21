using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;

namespace SharpLayout
{
    /// <summary>
    /// https://en.wikipedia.org/wiki/Option_type
    /// Java https://docs.oracle.com/javase/8/docs/api/java/util/Optional.html
    /// F# https://msdn.microsoft.com/en-us/visualfsharpdocs/conceptual/options-%5Bfsharp%5D
    /// C# Roslyn https://github.com/dotnet/roslyn/blob/56f605c41915317ccdb925f66974ee52282609e7/src/Compilers/Core/Portable/Optional.cs
    /// </summary>
    public struct Option<T>
    {
        private readonly T value;
        public bool HasValue { get; }

        public Option(T value)
        {
            this.value = value;
            HasValue = true;
        }

        public T Value
        {
            get
            {
                if (HasValue) return value;
                throw new InvalidOperationException($"Optional value of '{typeof (T)}' type has no value.");
            }
        }

        public TResult Match<TResult>(Func<T, TResult> some, Func<TResult> none) => HasValue ? some(Value) : none();

        public void Match(Action<T> some, Action none)
        {
            if (HasValue) some(Value);
            else none();
        }

        public T ValueOrDefault() => HasValue ? Value : default(T);

        [Pure]
        public T ValueOr(T defaultValue) => HasValue ? value : defaultValue;

        public override string ToString()
        {
            return HasValue ? Value.ToString() : "";
        }

        public static implicit operator Option<T>(T value) => new Option<T>(value);

        public Option<TResult> Select<TResult>(Func<T, TResult> func) 
            => HasValue ? func(Value) : new Option<TResult>();

        public Option<TResult> SelectMany<TResult>(Func<T, Option<TResult>> func) 
            => HasValue ? func(Value) : new Option<TResult>();

        public Option<TResult> SelectMany<TOption, TResult>(Func<T, Option<TOption>> optionFunc, Func<T, TOption, TResult> resultFunc)
            => SelectMany(value1 => optionFunc(value1).Select(value2 => resultFunc(value1, value2)));

        public Option<T> Where(Func<T, bool> predicate) 
            => HasValue ? (predicate(Value) ? this : new Option<T>()) : new Option<T>();
    }

    public static class Option
    {
        public static Option<T> AsOption<T>(this T it) => new Option<T>(it);

        public static Option<T> ToOption<T>(this T it) where T : class => it ?? new Option<T>();

        public static Option<T> ToOption<T>(this T? it) where T : struct => it ?? new Option<T>();

        public static Option<T> None<T>(this T it, Func<T, bool> predicate) => predicate(it) ? new Option<T>() : it;

        public static Option<TValue> Get<TKey, TValue>(this Dictionary<TKey, TValue> it, TKey key)
        {
            TValue value;
            return it.TryGetValue(key, out value) ? value : new Option<TValue>();
        }

        public static Option<T> FirstOrNone<T>(this IEnumerable<T> it)
        {
            var firstOrDefault = it.Select(_ => new {_}).FirstOrDefault();
            return firstOrDefault == null ? new Option<T>() : firstOrDefault._;
        }

        public static Option<T> LastOrNone<T>(this IEnumerable<T> it)
        {
            var lastOrDefault = it.Select(_ => new {_}).LastOrDefault();
            return lastOrDefault == null ? new Option<T>() : lastOrDefault._;
        }
    }
}