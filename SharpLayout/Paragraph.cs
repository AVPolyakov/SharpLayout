using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Paragraph : IElement
    {
        public List<Span> Spans { get; } = new List<Span>();

        private Option<double> leftMargin;
        public Option<double> LeftMargin() => leftMargin;
        public Paragraph LeftMargin(Option<double> value)
        {
            leftMargin = value;
            return this;
        }

        private Option<double> rightMargin;
        public Option<double> RightMargin() => rightMargin;
        public Paragraph RightMargin(Option<double> value)
        {
            rightMargin = value;
            return this;
        }

        private Option<double> topMargin;
        public Option<double> TopMargin() => topMargin;
        public Paragraph TopMargin(Option<double> value)
        {
            topMargin = value;
            return this;
        }

        private Option<double> bottomMargin;
        public Option<double> BottomMargin() => bottomMargin;
        public Paragraph BottomMargin(Option<double> value)
        {
            bottomMargin = value;
            return this;
        }

        private Option<double> textIndent;
        public Paragraph TextIndent(Option<double> value)
        {
            textIndent = value;
            return this;
        }

        public Option<double> TextIndent() => textIndent;

        private Func<double, double> lineSpacingFunc = _ => _;

        public Paragraph LineSpacingFunc(Func<double, double> value)
        {
            lineSpacingFunc = value;
            return this;
        }

        public Func<double, double> LineSpacingFunc() => lineSpacingFunc;

        public Paragraph Margin(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftMargin(value);
            if (direction.HasFlag(Right)) RightMargin(value);
            if (direction.HasFlag(Top)) TopMargin(value);
            if (direction.HasFlag(Bottom)) BottomMargin(value);
            return this;
        }

        private HorizontalAlign alignment;

        public Paragraph Alignment(HorizontalAlign value)
        {
            alignment = value;
            return this;
        }

        public HorizontalAlign Alignment() => alignment;

        public Paragraph Add(Span span, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            if (Spans.Count == 0)
                CallerInfos.Add(new CallerInfo {Line = line, FilePath = filePath});
            Spans.Add(span);
            return this;
        }

        public Paragraph Add(string text, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
            => Add(new Span(text, font), line, filePath);

	    public Paragraph Add(Expression<Func<string>> expression, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(new Span(expression, font), line, filePath);

	    public Paragraph Add<T>(Expression<Func<T>> expression, Func<T, string> converter, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(Span.Create(expression, converter, font), line, filePath);

        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table) => paragraph(this);

        public readonly List<CallerInfo> CallerInfos = new List<CallerInfo>();

        private bool keepWithNext;
        public bool KeepWithNext() => keepWithNext;
        public Paragraph KeepWithNext(bool value)
        {
            keepWithNext = value;
            return this;
        }
    }
}