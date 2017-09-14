using System;
using System.Collections.Generic;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Paragraph : IElement
    {
        public List<Span> Spans { get; } = new List<Span>();

        public Option<double> LeftMargin { get; set; }

        public Option<double> RightMargin { get; set; }

        public Option<double> TopMargin { get; set; }

        public Option<double> BottomMargin { get; set; }

        public Paragraph Margin(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftMargin = value;
            if (direction.HasFlag(Right)) RightMargin = value;
            if (direction.HasFlag(Top)) TopMargin = value;
            if (direction.HasFlag(Bottom)) BottomMargin = value;
            return this;
        }

        private HorizontalAlign alignment;

        public Paragraph Alignment(HorizontalAlign value)
        {
            alignment = value;
            return this;
        }

        public HorizontalAlign Alignment() => alignment;

        public Paragraph Add(Span span)
        {
            Spans.Add(span);
            return this;
        }

        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table) => paragraph(this);
    }
}