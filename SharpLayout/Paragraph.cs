using System.Collections.Generic;

namespace SharpLayout
{
    public class Paragraph
    {
        public List<Span> Spans { get; } = new List<Span>();

        public Option<double> LeftMargin { get; set; }

        public Option<double> RightMargin { get; set; }

        public Option<double> TopMargin { get; set; }

        public Option<double> BottomMargin { get; set; }

        public ParagraphAlignment Alignment { get; set; } = ParagraphAlignment.Left;

        public Paragraph Add(Span span)
        {
            Spans.Add(span);
            return this;
        }
    }
}