using System;

namespace SharpLayout
{
    public interface IElement
    {
        T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table, Func<Image, T> image);
    }
}