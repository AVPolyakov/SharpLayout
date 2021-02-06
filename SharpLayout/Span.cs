using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static SharpLayout.InlineVerticalAlign;

namespace SharpLayout
{
    public class Span
    {
        private readonly IValue textValue;
        private string GetTextOrEmpty(Document document, TextMode mode, Option<Table> table) => 
            Text(table).GetText(document, mode) ?? "";

        public ISoftLinePart[] GetSoftLineParts(Span span, Document document, DrawCache drawCache, Option<Table> table, TextMode mode)
        {
            var lines = GetTextOrEmpty(document, mode, table).SplitToLines();
            if (lines.Length == 0)
                return new ISoftLinePart[] {
                    new SoftLinePart(span, GetTextOrEmpty(document, mode, table),
                        drawCache.GetCharSizeCache(span.Font(table)))
                };
            else
            {
                var result = new ISoftLinePart[lines.Length];
                for (var index = 0; index < lines.Length; index++)
                    result[index] = new SoftLinePart(span, lines[index], drawCache.GetCharSizeCache(span.Font(table)));
                return result;
            }
        }

        private class SoftLinePart : ISoftLinePart
        {
            private readonly string stringText;
            public Span Span { get; }
            public CharSizeCache CharSizeCache { get; }

            public string Text(TextMode mode) => stringText;

            public IValue SubText(int startIndex, int length)
            {
                return new TextValue(stringText.Substring(startIndex, length));
            }

            public SoftLinePart(Span span, string stringText, CharSizeCache charSizeCache)
            {
                Span = span;
                CharSizeCache = charSizeCache;
                this.stringText = stringText;
            }
        }


        public IValue Text(Option<Table> table)
	    {
		    if (!FontOrNone(table).HasValue) return new TextValue("Font not set");
		    return textValue;
	    }

	    private Option<Font> font;
	    public Option<Font> Font() => font;
	    public Span Font(Option<Font> value)
	    {
		    font = value;
		    return this;
	    }

	    internal Option<Font> FontOrNone(Option<Table> table)
	    {
		    if (Font().HasValue) return Font().Value;
	        var tableFont = table.Select(_ => _.Font());
	        if (tableFont.HasValue) return tableFont.Value;
		    return new Option<Font>();
	    }

        internal Font Font(Option<Table> table)
        {
            var xFont = FontWithoutInlineVerticalAlign(table);
            if (InlineVerticalAlign() == Sub || InlineVerticalAlign() == Super)
            {
                var ascent = xFont.FontFamily.GetCellAscent(xFont.Style);
                var lineSpacing = xFont.FontFamily.GetLineSpacing(xFont.Style);
                var inlineVerticalAlignScaling = 0.8 * ascent / lineSpacing;
                return new Font(xFont.FamilyInfo, inlineVerticalAlignScaling * xFont.Size, xFont.Style, xFont.PdfOptions);
            }
            else
                return xFont;
        }

        internal Font FontWithoutInlineVerticalAlign(Option<Table> table) => FontOrNone(table).Match(
	        _ => _,
	        () => new Font(DefaultFontFamilies.Roboto, 10, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode)));


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

	    public Span(IValue value)
	    {
		    textValue = value;
	    }

	    public Span(string text): this(new TextValue(text))
	    {
	    }

	    public Span(Func<RenderContext, string> func) : this(new RenderContextValue(func))
	    {
	    }

	    public Span(Func<string> expression): 
		    this(ExpressionValue.Get(expression))
	    {
	    }

	    public static Span Create<T>(Func<T> expression, Func<T, string> converter) => 
		    new Span(ExpressionValue.Get(expression, converter));

	    public Span(IValue value, Font font) : this(value)
        {
            this.font = font;
        }

        public Span(string text, Font font): this(new TextValue(text), font)
        {
        }

        public Span(Func<RenderContext, string> func, Font font) : this(new RenderContextValue(func), font)
        {
        }

	    public Span(Func<string> expression, Font font): 
			this(ExpressionValue.Get(expression), font)
	    {
	    }

	    public static Span Create<T>(Func<T> expression, Func<T, string> converter, Font font) => 
		    new Span(ExpressionValue.Get(expression, converter), font);

        internal List<Func<Section, Table>> Footnotes { get; } = new List<Func<Section, Table>>();

        public Span AddFootnote(Table table)
        {
            Footnotes.Add(section => table);
            return this;
        }

	    public Span AddFootnote(Paragraph paragraph, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
	    {
	        Footnotes.Add(section => {
	            var table = new Table(line);
	            var c = table.AddColumn(section.PageSettings.PageWidthWithoutMargins);
	            var r = table.AddRow();
	            r[c, line, filePath].Add(paragraph);
	            return table;
	        });
            return this;
	    }
    }

    public interface IValue
	{
		string GetText(Document document, TextMode mode);
		bool IsExpression { get; }
	}

	public class TextValue : IValue
	{
		private readonly string text;

		public string GetText(Document document, TextMode mode) => text;

		public bool IsExpression => false;

		public TextValue(string text)
		{
			this.text = text;
		}
	}

    public class RenderContextValue : IValue
    {
        private readonly Func<RenderContext, string> func;

        public string GetText(Document document, TextMode mode) => func(mode.RenderContext);

        public bool IsExpression => false;

        public RenderContextValue(Func<RenderContext, string> func)
        {
            this.func = func;
        }
    }

    public class RenderContext
    {
        private readonly int pageIndex;

        public RenderContext(int pageIndex, int pagesCount)
        {
            this.pageIndex = pageIndex;
            PageCount = pagesCount;
        }

        public int PageNumber => pageIndex + 1;

        public int PageCount { get; }
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

		public string GetText(Document document, TextMode mode) => 
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

    public interface ISoftLinePart
    {
        Span Span { get; }
        CharSizeCache CharSizeCache { get; }
        string Text(TextMode mode);
        IValue SubText(int startIndex, int length);
    }

    public abstract class TextMode
    {
	    public const int DefaultNumber = 8;
	    
        public class Measure: TextMode
        {
	        public override RenderContext RenderContext => new RenderContext(pageIndex: DefaultNumber, pagesCount: DefaultNumber);
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

            public override RenderContext RenderContext => new RenderContext(PageIndex, PagesCount);
        }
        
        public abstract RenderContext RenderContext { get; }
    }
}