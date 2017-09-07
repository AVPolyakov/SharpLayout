using System;
using System.Collections.Generic;

namespace SharpLayout
{
    public class Paragraph : IElement
    {
        public List<Span> Spans { get; } = new List<Span>();

        public Option<double> LeftMargin { get; set; }

        public Option<double> RightMargin { get; set; }

        public Option<double> TopMargin { get; set; }

        public Option<double> BottomMargin { get; set; }

        public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.Left;

        public Paragraph Add(Span span)
        {
            Spans.Add(span);
            return this;
        }

        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table) => paragraph(this);
    }
}