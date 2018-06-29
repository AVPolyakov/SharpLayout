using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Paragraph : IElement
    {
        public List<Span> Spans { get; } = new List<Span>();
        
        internal bool IsParagraphPart { get; set; }

        private double? leftMargin;
        public double? LeftMargin() => leftMargin;
        public Paragraph LeftMargin(double? value)
        {
            leftMargin = value;
            return this;
        }

        private double? rightMargin;
        public double? RightMargin() => rightMargin;
        public Paragraph RightMargin(double? value)
        {
            rightMargin = value;
            return this;
        }

        private double? topMargin;
        public double? TopMargin() => topMargin;
        public Paragraph TopMargin(double? value)
        {
            topMargin = value;
            return this;
        }

        private double? bottomMargin;
        public double? BottomMargin() => bottomMargin;
        public Paragraph BottomMargin(double? value)
        {
            bottomMargin = value;
            return this;
        }

        private double? textIndent;
        public Paragraph TextIndent(double? value)
        {
            textIndent = value;
            return this;
        }

        public double? TextIndent() => textIndent;

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

        private HorizontalAlign? alignment;
	    public HorizontalAlign? Alignment() => alignment;
        public Paragraph Alignment(HorizontalAlign? value)
        {
            alignment = value;
            return this;
        }

        public Paragraph Add(Span span, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            if (Spans.Count == 0)
                CallerInfos.Add(new CallerInfo {Line = line, FilePath = filePath});
            Spans.Add(span);
            return this;
        }

	    public Paragraph Add(string text, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(new Span(text), line, filePath);

	    public Paragraph Add(Func<string> expression, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(new Span(expression), line, filePath);

	    public Paragraph Add<T>(Func<T> expression, Func<T, string> converter, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(Span.Create(expression, converter), line, filePath);

        public Paragraph Add(string text, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
            => Add(new Span(text, font), line, filePath);

	    public Paragraph Add(Func<string> expression, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(new Span(expression, font), line, filePath);

	    public Paragraph Add<T>(Func<T> expression, Func<T, string> converter, XFont font, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "") 
		    => Add(Span.Create(expression, converter, font), line, filePath);

        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table) => paragraph(this);

        public readonly List<CallerInfo> CallerInfos = new List<CallerInfo>();

        private bool? keepWithNext;
        public bool? KeepWithNext() => keepWithNext;
        public Paragraph KeepWithNext(bool? value)
        {
            keepWithNext = value;
            return this;
        }
    }
}