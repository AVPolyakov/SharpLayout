using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static SharpLayout.InlineVerticalAlign;

namespace SharpLayout
{
    public class Span
    {
	    private readonly IText text;

	    public IText Text(Option<Table> table)
	    {
		    if (!FontOrNone(table).HasValue) return new Text(new TextValue("Font not set"));
		    return text;
	    }

	    private Option<XFont> font;
	    public Option<XFont> Font() => font;
	    public Span Font(Option<XFont> value)
	    {
		    font = value;
		    return this;
	    }

	    internal Option<XFont> FontOrNone(Option<Table> table)
	    {
		    if (Font().HasValue) return Font().Value;
	        var tableFont = table.Select(_ => _.Font());
	        if (tableFont.HasValue) return tableFont.Value;
		    return new Option<XFont>();
	    }

        internal XFont Font(Option<Table> table)
        {
            var xFont = FontWithoutInlineVerticalAlign(table);
            if (InlineVerticalAlign() == Sub || InlineVerticalAlign() == Super)
            {
                var ascent = xFont.FontFamily.GetCellAscent(xFont.Style);
                var lineSpacing = xFont.FontFamily.GetLineSpacing(xFont.Style);
                var inlineVerticalAlignScaling = 0.8 * ascent / lineSpacing;
                return new XFont(xFont.Name, inlineVerticalAlignScaling * xFont.Size, xFont.Style, xFont.PdfOptions);
            }
            else
                return xFont;
        }

        internal XFont FontWithoutInlineVerticalAlign(Option<Table> table) => FontOrNone(table).ValueOr(defaultFont);

        private static readonly XFont defaultFont = new XFont("Times New Roman", 10, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode));

        private InlineVerticalAlign inlineVerticalAlign;
        public InlineVerticalAlign InlineVerticalAlign() => inlineVerticalAlign;
        public Span InlineVerticalAlign(InlineVerticalAlign value)
        {
            inlineVerticalAlign = value;
            return this;
        }

        private Option<XBrush> brush;
        public Option<XBrush> Brush() => brush;
        public Span Brush(Option<XBrush> value)
        {
            brush = value;
            return this;
        }

        private XColor? backgroundColor;
        public XColor? BackgroundColor() => backgroundColor;
        public Span BackgroundColor(XColor? value)
        {
            backgroundColor = value;
            return this;
        }

	    public Span(IText text)
	    {
		    this.text = text;
	    }

	    public Span(string text): this(new Text(new TextValue(text)))
	    {
	    }

	    public Span(Func<string> expression): 
		    this(new Text(ExpressionValue.Get(expression)))
	    {
	    }

	    public static Span Create<T>(Func<T> expression, Func<T, string> converter) => 
		    new Span(new Text(ExpressionValue.Get(expression, converter)));

	    public Span(IText text, XFont font) : this(text)
        {
            this.font = font;
        }

        public Span(string text, XFont font): this(new Text(new TextValue(text)), font)
        {
        }

	    public Span(Func<string> expression, XFont font): 
			this(new Text(ExpressionValue.Get(expression)), font)
	    {
	    }

	    public static Span Create<T>(Func<T> expression, Func<T, string> converter, XFont font) => 
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
		private readonly Func<T> expression;
		private readonly Func<T, string> converter;

		public ExpressionValue(Func<T> expression, Func<T, string> converter)
		{
			this.expression = expression;
			this.converter = converter;
		}

		public string GetText(Document document) => 
			document.ExpressionVisible ? ReflectionUtil.GetMemberInfo(expression).Name : converter(expression());

		public bool IsExpression => true;
	}

	public static class ExpressionValue
	{
		public static ExpressionValue<T> Get<T>(Func<T> expression, Func<T, string> converter) => 
			new ExpressionValue<T>(expression, converter);

		public static ExpressionValue<string> Get(Func<string> expression) => 
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

            public IText SubText(int startIndex, int length)
            {
                return new Text(new TextValue(stringText.Substring(startIndex, length)));
            }

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

            public IText SubText(int startIndex, int length)
            {
                return new PageNumber();
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

            public IText SubText(int startIndex, int length)
            {
                return new PageCount();
            }
        }

	    public bool IsExpression => false;
    }

    public interface ISoftLinePart
    {
        Span Span { get; }
        string Text(TextMode mode);
        IText SubText(int startIndex, int length);
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