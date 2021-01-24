using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using static SharpLayout.ParagraphRenderer;

namespace SharpLayout
{
    public class Section
    {
        public PageSettings PageSettings { get; }
        
        internal readonly List<List<Func<Document, IGraphics, Table[]>>> tableFuncs = 
            new List<List<Func<Document, IGraphics, Table[]>>>{new List<Func<Document, IGraphics, Table[]>>()};
        
        public List<Table> Headers { get; } = new List<Table>();
        public Option<List<Table>> FirstPageHeaders { get; set; }
        public Option<List<Table>> EvenPageHeaders  { get; set; }
        
        public List<Table> Footers { get; } = new List<Table>();
        public Option<List<Table>> FirstPageFooters { get; set; }
        public Option<List<Table>> EvenPageFooters { get; set; }
        
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
            LastTableFuncs.Add((d, g) => new[] {table});
            return this;
        }

        private List<Func<Document, IGraphics, Table[]>> LastTableFuncs => tableFuncs[tableFuncs.Count - 1];

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

        public Section SetEmptyFirstPageHeaders()
        {
	        FirstPageHeaders = new List<Table>();
	        return this;
        }
        
        public Table AddFirstPageHeader([CallerLineNumber] int line = 0)
        {
	        var table = new Table(line);
	        AddFirstPageHeader(table);
	        return table;
        }
        
        public Section AddFirstPageHeader(Table table)
        {
	        List<Table> headers;
	        if (!FirstPageHeaders.HasValue)
	        {
		        headers = new List<Table>();
		        FirstPageHeaders = headers;
	        }
	        else
		        headers = FirstPageHeaders.Value;
	        headers.Add(table);
	        return this;
        }

        public Section SetEmptyEvenPageHeaders()
        {
	        EvenPageHeaders = new List<Table>();
	        return this;
        }
        
        public Table AddEvenPageHeader([CallerLineNumber] int line = 0)
        {
	        var table = new Table(line);
	        AddEvenPageHeader(table);
	        return table;
        }
        
        public Section AddEvenPageHeader(Table table)
        {
	        List<Table> headers;
	        if (!EvenPageHeaders.HasValue)
	        {
		        headers = new List<Table>();
		        EvenPageHeaders = headers;
	        }
	        else
		        headers = EvenPageHeaders.Value;
	        headers.Add(table);
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

        public Section SetEmptyFirstPageFooters()
        {
	        FirstPageFooters = new List<Table>();
	        return this;
        }
        
        public Table AddFirstPageFooter([CallerLineNumber] int line = 0)
        {
	        var table = new Table(line);
	        AddFirstPageFooter(table);
	        return table;
        }
        
        public Section AddFirstPageFooter(Table table)
        {
	        List<Table> footers;
	        if (!FirstPageFooters.HasValue)
	        {
		        footers = new List<Table>();
		        FirstPageFooters = footers;
	        }
	        else
		        footers = FirstPageFooters.Value;
	        footers.Add(table);
	        return this;
        }

        public Section SetEmptyEvenPageFooters()
        {
	        EvenPageFooters = new List<Table>();
	        return this;
        }
        
        public Table AddEvenPageFooter([CallerLineNumber] int line = 0)
        {
	        var table = new Table(line);
	        AddEvenPageFooter(table);
	        return table;
        }
        
        public Section AddEvenPageFooter(Table table)
        {
	        List<Table> footers;
	        if (!EvenPageFooters.HasValue)
	        {
		        footers = new List<Table>();
		        EvenPageFooters = footers;
	        }
	        else
		        footers = EvenPageFooters.Value;
	        footers.Add(table);
	        return this;
        }
        
        public Section AddPageBreak()
        {
            tableFuncs.Add(new List<Func<Document, IGraphics, Table[]>>());
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
            if (paragraph.KeepWithNext().GetValueOrDefault(false) || 
                paragraph.KeepLinesTogether().GetValueOrDefault(false))
            {
                var table = new Table(line).KeepWithNext(paragraph.KeepWithNext());
                LastTableFuncs.Add((document, graphics) => new []{table});
                var c1 = table.AddColumn(PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1, line, filePath].Add(paragraph);
            }
            else
            {
                var width = PageSettings.PageWidthWithoutMargins;
                var drawCache = new DrawCache();
                LastTableFuncs.Add((document, graphics) => {
	                var index = 0;
                    var mode = new TextMode.Measure();
                    var lines = paragraph.GetSoftLines(document, new Option<Table>(), drawCache, mode)
		                .SelectMany(softLineParts => {
			                var charInfos = GetCharInfos(softLineParts, mode);
			                return GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width),
					                charInfos, paragraph, mode, document, new Option<Table>())
				                .Select(lineInfo => {
					                var lineParts = lineInfo.GetLineParts(charInfos);
					                return lineParts.Count > 0
						                ? lineParts.Select(linePart => new {
								                linePart.GetSoftLinePart(softLineParts).Span,
								                Text = linePart.SubText(softLineParts)
							                }
						                )
						                : softLineParts.Select(softLinePart => new {
								                softLinePart.Span,
								                Text = (IValue) new TextValue("")
							                }
						                );
				                });
		                })
		                .Select(enumerable => enumerable.Select(_ => new {_.Span, _.Text, index = index++}).ToList())
		                .ToList();
	                var dictionary = lines.SelectMany(_ => _)
		                .ToLookup(_ => _.Span)
		                .Select(_ => new {_.Key, MaxIndex = _.Max(x => x.index)})
		                .ToDictionary(_ => _.Key, _ => _.MaxIndex);
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
                        p.Spans.AddRange(spans.Select(_ => Clone(_.Span, _.Text, _.index == dictionary[_.Span])));
                        p.CallerInfos?.AddRange(paragraph.CallerInfos);
                        var table = new Table(line).KeepWithNext(p.KeepWithNext());
                        var c1 = table.AddColumn(width);
                        var r1 = table.AddRow();
                        r1[c1, line, filePath].Add(p);
                        return table;
                    }).ToArray();
                });
            }
            return this;
        }

	    private static Span Clone(Span span, IValue subText, bool isLast)
	    {
		    var clone = new Span(subText)
			    .Font(span.Font())
			    .Brush(span.Brush())
			    .InlineVerticalAlign(span.InlineVerticalAlign())
			    .BackgroundColor(span.BackgroundColor());
		    if (isLast)
			    foreach (var footnote in span.Footnotes)
				    clone.Footnotes.Add(footnote);
		    return clone;
	    }
	    
	    private Option<Func<Stream>> template;
	    public Option<Func<Stream>> Template() => template;
	    public Section Template(Func<Stream> value)
	    {
		    template = value.ToOption();
		    return this;
	    }
    }
}