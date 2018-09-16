using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using static SharpLayout.ParagraphRenderer;

namespace SharpLayout
{
    public class Section
    {
        public PageSettings PageSettings { get; }
        private readonly List<Func<Document, XGraphics, Table[]>> tableFuncs = new List<Func<Document, XGraphics, Table[]>>();
        public List<Table> Headers { get; } = new List<Table>();
        public List<Table> Footers { get; } = new List<Table>();
        public List<Table> FootnoteSeparators { get; } = new List<Table>();

        public Section(PageSettings pageSettings)
        {
            PageSettings = pageSettings;
        }

        public Table AddTable([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            Add(table);
            return table;
        }

        public Section Add(Table table)
        {
            tableFuncs.Add((d, g) => new[] {table});
            return this;
        }

        public Table AddHeader([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            AddHeader(table);
            return table;
        }

        public Section AddHeader(Table table)
        {
            Headers.Add(table);
            return this;
        }

        public Table AddFooter([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            AddFooter(table);
            return table;
        }

        public Section AddFooter(Table table)
        {
            Footers.Add(table);
            return this;
        }

        /// <summary>
        /// See https://superuser.com/q/1130373
        /// </summary>
        public Section AddFootnoteSeparator(Table table)
        {
            FootnoteSeparators.Add(table);
            return this;
        }

        public Section Add(Paragraph paragraph, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            if (paragraph.KeepWithNext().GetValueOrDefault(false))
            {
                var table = new Table(paragraph.KeepWithNext(), line);
                tableFuncs.Add((document, graphics) => new []{table});
                var c1 = table.AddColumn(PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1, line, filePath].Add(paragraph);
            }
            else
            {
                var width = PageSettings.PageWidthWithoutMargins;
                tableFuncs.Add((document, graphics) => {
                    var lines = paragraph.GetSoftLines(document, new Option<Table>())
                        .SelectMany(softLineParts => {
                            var charInfos = GetCharInfos(softLineParts, new TextMode.Measure());
                            return GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width),
                                    charInfos, paragraph, new TextMode.Measure(), document, new Option<Table>())
                                .Select(lineInfo => {
                                    var lineParts = lineInfo.GetLineParts(charInfos).ToList();
                                    return lineParts.Any()
			                            ? lineParts.Select((linePart, i) => Clone(
				                            linePart.GetSoftLinePart(softLineParts).Span,
				                            linePart.SubText(softLineParts), i == lineParts.Count - 1))
			                            : softLineParts.Select(softLinePart => Clone(
				                            softLinePart.Span,
				                            new Text(new TextValue("")), true));
	                            });
                        }).ToList();
                    return lines.Select((spans, i) => {
                        var p = new Paragraph {IsParagraphPart = i < lines.Count - 1}
                            .Alignment(paragraph.Alignment())
                            .LeftMargin(paragraph.LeftMargin())
                            .RightMargin(paragraph.RightMargin())                            
                            .LineSpacingFunc(paragraph.LineSpacingFunc())
                            .KeepWithNext(i == lines.Count - 2);
                        if (i == 0)
                            p.TextIndent(paragraph.TextIndent())
                                .TopMargin(paragraph.TopMargin());
                        if (i == lines.Count - 1)
                            p.BottomMargin(paragraph.BottomMargin());
                        p.Spans.AddRange(spans);
                        p.CallerInfos?.AddRange(paragraph.CallerInfos);
                        var table = new Table(p.KeepWithNext(), line);
                        var c1 = table.AddColumn(width);
                        var r1 = table.AddRow();
                        r1[c1, line, filePath].Add(p);
                        return table;
                    }).ToArray();
                });
            }
            return this;
        }

	    private static Span Clone(Span span, IText subText, bool last)
	    {
	        var clone = new Span(subText)
	            .Font(span.Font())
	            .Brush(span.Brush())
	            .InlineVerticalAlign(span.InlineVerticalAlign())
	            .BackgroundColor(span.BackgroundColor());
	        if (last)
	            foreach (var footnote in span.Footnotes)
	                clone.AddFootnote(footnote);
	        return clone;
	    }

	    public List<Table> GetTables(Document document, XGraphics xGraphics)
        {
            return tableFuncs.Select(func => func(document, xGraphics)).SelectMany(_ => _).ToList();
        }
    }
}