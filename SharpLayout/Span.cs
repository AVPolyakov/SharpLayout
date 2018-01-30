using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Span
    {
        public IText Text { get; }
        public XFont Font { get; }
        
        public XBrush Brush { get; set; } = XBrushes.Black;

        private Option<XColor> backgroundColor;

        public Span BackgroundColor(XColor value)
        {
            backgroundColor = value;
            return this;
        }

        public Option<XColor> BackgroundColor() => backgroundColor;

	    public Span(IText text, XFont font)
        {
            Text = text;
            Font = font;
        }

        public Span(string text, XFont font): this(new Text(new TextValue(text)), font)
        {
        }

	    public Span(Expression<Func<string>> expression, XFont font): 
			this(new Text(ExpressionValue.Get(expression)), font)
	    {
	    }

	    public static Span Create<T>(Expression<Func<T>> expression, Func<T, string> converter, XFont font) => 
		    new Span(new Text(ExpressionValue.Get(expression, converter)), font);
    }

	public interface IValue
	{
		string GetText(Document document);
		bool IsExpression { get; }
	}

	public class TextValue : IValue
	{
		private readonly string text;

		public string GetText(Document document) => text;

		public bool IsExpression => false;

		public TextValue(string text)
		{
			this.text = text;
		}
	}

	public class ExpressionValue<T> : IValue
	{
		private readonly Expression<Func<T>> expression;
		private readonly Func<T, string> converter;

		public ExpressionValue(Expression<Func<T>> expression, Func<T, string> converter)
		{
			this.expression = expression;
			this.converter = converter;
		}

		public string GetText(Document document) => 
			document.ExpressionVisible ? ((MemberExpression) expression.Body).Member.Name : converter(expression.Compile()());

		public bool IsExpression => true;
	}

	public static class ExpressionValue
	{
		public static ExpressionValue<T> Get<T>(Expression<Func<T>> expression, Func<T, string> converter) => 
			new ExpressionValue<T>(expression, converter);

		public static ExpressionValue<string> Get(Expression<Func<string>> expression) => 
			Get(expression, StringConverter);

		public static Func<string, string> StringConverter => _ => _;
	}

    public interface IText
    {
        IEnumerable<ISoftLinePart> GetSoftLineParts(Span span, Document document);
	    bool IsExpression { get; }
    }

    public class Text : IText
    {
	    private readonly IValue value;
	    private string GetTextOrEmpty(Document document) => value.GetText(document) ?? "";

	    public Text(IValue value)
        {
	        this.value = value;
        }

        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span, Document document)
        {
            var lines = GetTextOrEmpty(document).SplitToLines();
            if (!lines.Any())
                yield return new SoftLinePart(span, GetTextOrEmpty(document));
            else
                foreach (var line in lines)
                    yield return new SoftLinePart(span, line);
        }

        private class SoftLinePart: ISoftLinePart
        {
            private readonly string stringText;
	        public Span Span { get; }

            public string Text(TextMode mode) => stringText;

	        public SoftLinePart(Span span, string stringText)
            {
                Span = span;
                this.stringText = stringText;
            }
        }

	    public bool IsExpression => value.IsExpression;
    }

    public class PageNumber : IText
    {
        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span, Document document)
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

	    public bool IsExpression => false;
    }

    public class PageCount : IText
    {
        public IEnumerable<ISoftLinePart> GetSoftLineParts(Span span, Document document)
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

	    public bool IsExpression => false;
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