using System;
using SharpLayout;

namespace Watcher
{
    public abstract class Choice<T1, T2>
    {
        public abstract T Match<T>(Func<T1, T> choice1, Func<T2, T> choice2);

        public static implicit operator Choice<T1, T2>(T1 choice1) => new Choice1(choice1);

        public static implicit operator Choice<T1, T2>(T2 choice2) => new Choice2(choice2);

        public bool HasValue1 => Match(_ => true, _ => false);

        public bool HasValue2 => Match(_ => false, _ => true);

        public T1 Value1 => Match(_ => _, _ => throw new InvalidOperationException());

        public T2 Value2 => Match(_ => throw new InvalidOperationException(), _ => _);

        private Choice()
        {
        }

        public sealed class Choice1 : Choice<T1, T2>
        {
            private readonly T1 value;

            public Choice1(T1 value) => this.value = value;

            public override T Match<T>(Func<T1, T> choice1, Func<T2, T> choice2) => choice1(value);
        }

        public sealed class Choice2 : Choice<T1, T2>
        {
            private readonly T2 value;

            public Choice2(T2 value) => this.value = value;

            public override T Match<T>(Func<T1, T> choice1, Func<T2, T> choice2) => choice2(value);
        }
    }
    
    public static class Choice
    {
        public static Choice<T1, T2> Create<T1, T2>(T1 choice1, Option<T2> choice2) => choice1;

        public static Choice<T1, T2> Create<T1, T2>(Option<T1> choice1, T2 choice2) => choice2;
    }
}