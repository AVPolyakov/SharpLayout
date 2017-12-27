using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Span
    {
        private readonly IText text;
        public XFont Font { get; }
        
        public XBrush Brush { get; set; } = XBrushes.Black;

        private Option<XColor> backgroundColor;

        public Span BackgroundColor(XColor value)
        {
            backgroundColor = value;
            return this;
        }

        public Option<XColor> BackgroundColor() => backgroundColor;

        public IEnumerable<ISoftLinePart> GetSoftLineParts()
        {
            return text.GetSoftLineParts(this);
        }

        public Span(IText text, XFont font)
        {
            this.text = text;
            Font = font;
        }

        public Span(string text, XFont font): this(new Text(text), font)
        {
        }
    }

    public interface IText
    {
        IEnumerable<ISoftLinePart> GetSoftLineParts(Span span);
    }

    public class Text : IText
    {
        private readonly string text;
        private string TextOrEmpty => text ?? "";

        public Text(string text)
        {
            this.text = text;
        }

        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span)
        {
            var lines = TextOrEmpty.SplitToLines();
            if (!lines.Any())
                yield return new SoftLinePart(span, TextOrEmpty);
            else
                foreach (var line in lines)
                    yield return new SoftLinePart(span, line);
        }

        private class SoftLinePart: ISoftLinePart
        {
            private readonly string text;
            public Span Span { get; }

            public string Text(TextMode mode) => text;

            public SoftLinePart(Span span, string text)
            {
                Span = span;
                this.text = text;
            }
        }
    }

    public class PageNumber : IText
    {
        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span)
        {
            yield return new SoftLinePart(span);
        }

        private class SoftLinePart : ISoftLinePart
        {
            public SoftLinePart(Span span)
            {
                Span = span;
            }

            public Span Span { get; }

            public string Text(TextMode mode)
            {
                switch (mode)
                {
                    case TextMode.Measure _:
                        return "8";
                    case TextMode.Draw draw:
                        return $"{draw.PageIndex + 1}";
                    default:
                        throw new ArgumentOutOfRangeException(nameof(mode));
                }
            }
        }
    }

    public class PageCount : IText
    {
        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span)
        {
            yield return new SoftLinePart(span);
        }

        private class SoftLinePart : ISoftLinePart
        {
            public SoftLinePart(Span span)
            {
                Span = span;
            }

            public Span Span { get; }

            public string Text(TextMode mode)
            {
                switch (mode)
                {
                    case TextMode.Measure _:
                        return "8";
                    case TextMode.Draw draw:
                        return $"{draw.PagesCount}";
                    default:
                        throw new ArgumentOutOfRangeException(nameof(mode));
                }
            }
        }
    }

    public interface ISoftLinePart
    {
        Span Span { get; }
        string Text(TextMode mode);
    }

    public class TextMode
    {
        public class Measure: TextMode
        {            
        }

        public class Draw: TextMode
        {
            public int PageIndex { get; }
            public int PagesCount { get; }

            public Draw(int pageIndex, int pagesCount)
            {
                PageIndex = pageIndex;
                PagesCount = pagesCount;
            }
        }

        protected TextMode()
        {
        }
    }
}