using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public static class ParagraphRenderer
    {
	    public static void Draw(XGraphics graphics, Paragraph paragraph, XUnit x0, XUnit y0, double width, HorizontalAlign alignment, Drawer drawer,
		    GraphicsType graphicsType, TextMode mode, Document document, Table table)
        {
            var y = y0 + paragraph.TopMargin().ToOption().ValueOr(0);
            var lineCount = Lazy.Create(() => GetLineCount(graphics, paragraph, width, mode, document, table));
            var lineIndex = 0;
            double TextIndent() => lineIndex == 0 ? paragraph.TextIndent().ToOption().ValueOr(0) : 0;
            foreach (var softLineParts in paragraph.GetSoftLines(document, table))
            {
                var charInfos = GetCharInfos(softLineParts, mode);
                var innerWidth = paragraph.GetInnerWidth(width);
                var lineInfos = GetLines(graphics, softLineParts, innerWidth, charInfos, paragraph, mode, document, table).ToList();
                foreach (var line in lineInfos)
                {
                    var lineParts = line.GetLineParts(charInfos).ToList();
                    double dx;
                    switch (alignment)
                    {
                        case HorizontalAlign.Left:
                        case HorizontalAlign.Justify:
                            dx = 0;
                            break;
                        case HorizontalAlign.Center:
                            dx = (innerWidth - lineParts.ContentWidth(softLineParts, graphics, mode, table) - TextIndent()) / 2;
                            break;
                        case HorizontalAlign.Right:
                            dx = innerWidth - lineParts.ContentWidth(softLineParts, graphics, mode, table) - TextIndent();
                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
                    }
                    var baseLine = lineParts.Spans(softLineParts).Max(span => BaseLine(span.FontWithoutInlineVerticalAlign(table)));
                    var x = x0 + paragraph.LeftMargin().ToOption().ValueOr(0) + dx + TextIndent();
                    var maxLineSpace = lineParts.Spans(softLineParts).Max(span => span.FontWithoutInlineVerticalAlign(table).GetHeight());
                    var multiplier = Lazy.Create(() => {
                        var spaces = GetSpaces(lineParts, softLineParts, mode);
                        return spaces.Any()
                            ? (innerWidth - lineParts.ContentWidth(softLineParts, graphics, mode, table) - TextIndent()) /
                            spaces.Sum(tuple1 => graphics.MeasureString(new string(tuple1.Item2.Char, 1),
                                tuple1.Item1.GetSoftLinePart(softLineParts).Span.Font(table), MeasureTrailingSpacesStringFormat).Width)
                            : new double?();
                    });
                    foreach (var part in lineParts)
                    {
                        var text = part.Text(softLineParts, mode);
                        var span = part.GetSoftLinePart(softLineParts).Span;
                        var rectangleWidth = 0d;
                        var rectangleX = x;
	                    XFont font;
	                    if (alignment == HorizontalAlign.Justify && span.Font(table).Underline)
		                    font = new XFont(span.Font(table).FontFamily.Name, span.Font(table).Size,
			                    new[] {
				                    XFontStyle.Regular,
				                    XFontStyle.Bold,
				                    XFontStyle.Italic,
				                    XFontStyle.Strikeout,
			                    }.Where(_ => span.Font(table).Style.HasFlag(_)).Aggregate((style1, style2) => style1 | style2),
			                    span.Font(table).PdfOptions);
	                    else
		                    font = span.Font(table);
	                    if (alignment == HorizontalAlign.Justify || graphicsType == GraphicsType.Image && !font.Underline)
                            foreach (var drawTextPart in GetDrawTextParts(text))
                            {
                                double stringWidth;
                                switch (drawTextPart)
                                {
                                    case DrawTextPart.Space space:
                                        var measureWidth = graphics.MeasureString(new string(space.Char, 1), font, MeasureTrailingSpacesStringFormat).Width;
                                        if (multiplier.Value.HasValue)
                                            if (paragraph.IsParagraphPart || lineIndex < lineCount.Value - 1)
                                                if (alignment == HorizontalAlign.Justify)
                                                    stringWidth = measureWidth + measureWidth * multiplier.Value.Value;
                                                else
                                                    stringWidth = measureWidth;
                                            else
                                                stringWidth = measureWidth;
                                        else
                                            stringWidth = measureWidth;
                                        break;
                                    case DrawTextPart.Word word:
                                        drawer.DrawString(word.Text, font, span.CalculateBrush(document, table), x, y + baseLine + span.InlineVerticalAlignOffset(table));
                                        stringWidth = graphics.MeasureString(word.Text, font, MeasureTrailingSpacesStringFormat).Width;
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }
                                x += stringWidth;
                                rectangleWidth += stringWidth;
                            }
                        else
                        {
                            var measureString = graphics.MeasureString(text, font, MeasureTrailingSpacesStringFormat);
                            drawer.DrawString(text, font, span.CalculateBrush(document, table), x, y + baseLine + span.InlineVerticalAlignOffset(table));
                            x += measureString.Width;
                            rectangleWidth += measureString.Width;
                        }
                        if (span.BackgroundColor().HasValue)
                            drawer.DrawRectangle(new XSolidBrush(span.BackgroundColor().Value), rectangleX, y,
                                rectangleWidth,
                                maxLineSpace,
                                DrawType.Background);
	                    if (alignment == HorizontalAlign.Justify && span.Font(table).Underline)
		                    if (span.CalculateBrush(document, table) is XSolidBrush solidBrush)
		                    {
			                    var d = font.GetHeight() * (font.FontFamily.GetCellDescent(font.Style))
				                    / font.FontFamily.GetLineSpacing(font.Style);
			                    double yMultiplier;
			                    double widthMultiplier;
			                    if (font.Bold)
			                    {
				                    widthMultiplier = 0.096d;
				                    yMultiplier = 0.72509d;
			                    }
			                    else
			                    {
				                    yMultiplier = 0.61609d;
				                    widthMultiplier = 0.05d;
			                    }
			                    var lineY = y + baseLine + d * yMultiplier;
			                    drawer.DrawLine(new XPen(solidBrush.Color, font.Size*widthMultiplier),
				                    rectangleX, lineY, rectangleX + rectangleWidth, lineY);
		                    }
                    }
                    y += paragraph.LineSpacingFunc()(maxLineSpace);
                    lineIndex++;
                }
            }
        }

        private static double InlineVerticalAlignOffset(this Span span, Table table)
        {
            switch (span.InlineVerticalAlign())
            {
                case InlineVerticalAlign.Baseline:
                    return 0;
                case InlineVerticalAlign.Sub:
                    return -BaseLine(span.FontWithoutInlineVerticalAlign(table)) + BaseLine(span.Font(table)) +
                        span.FontWithoutInlineVerticalAlign(table).GetHeight() - span.Font(table).GetHeight();
                case InlineVerticalAlign.Super:
                    return -BaseLine(span.FontWithoutInlineVerticalAlign(table)) + BaseLine(span.Font(table));
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static XBrush CalculateBrush(this Span span, Document document, Table table)
	    {
		    if (!span.FontOrNone(table).HasValue) return new XSolidBrush(XColor.FromArgb(255, 0, 255));
		    if (span.Text(table).ExpressionVisible(document)) return XBrushes.Red;
		    return span.Brush().ValueOr(XBrushes.Black);
	    }

	    private static IEnumerable<Tuple<LinePart, DrawTextPart.Space>> GetSpaces(List<LinePart> lineParts, List<ISoftLinePart> softLineParts, TextMode mode)
        {
            foreach (var part in lineParts)
            foreach (var drawTextPart in GetDrawTextParts(part.Text(softLineParts, mode)))
                switch (drawTextPart)
                {
                    case DrawTextPart.Space space:
                        yield return Tuple.Create(part, space);
                        break;
                    case DrawTextPart.Word _:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
        }

        private static IEnumerable<DrawTextPart> GetDrawTextParts(string text)
        {
            var previousSpaceIndex = -1;
            for (var i = 0; i < text.Length; i++)
                if (char.IsWhiteSpace(text[i]))
                {
                    if (i - previousSpaceIndex - 1 > 0)
                        yield return new DrawTextPart.Word(text.Substring(previousSpaceIndex + 1, i - previousSpaceIndex - 1));
                    yield return new DrawTextPart.Space(text[i]);
                    previousSpaceIndex = i;
                }
            if (previousSpaceIndex + 1 < text.Length)
                yield return new DrawTextPart.Word(text.Substring(previousSpaceIndex + 1));
        }

        private class DrawTextPart
        {
            public class Word: DrawTextPart
            {
                public string Text { get; }

                public Word(string text)
                {
                    Text = text;
                }
            }

            public class Space: DrawTextPart
            {
                public char Char { get; }

                public Space(char @char)
                {
                    Char = @char;
                }
            }
        }

        public static double GetInnerWidth(this Paragraph paragraph, double width) 
            => width - paragraph.LeftMargin().ToOption().ValueOr(0) - paragraph.RightMargin().ToOption().ValueOr(0);

        internal static List<List<ISoftLinePart>> GetSoftLines(this Paragraph paragraph, Document document, Option<Table> table)
        {
            var result = new List<List<ISoftLinePart>>();
            foreach (var span in paragraph.Spans)
            {
                if (result.Count == 0)
                    result.Add(new List<ISoftLinePart>());
                var parts = span.Text(table).GetSoftLineParts(span, document).ToList();
                result[result.Count - 1].Add(parts[0]);
                for (var i = 1; i < parts.Count; i++)
                    result.Add(new List<ISoftLinePart> {parts[i]});
            }
            return result;
        }

        private static double ContentWidth(this List<LinePart> lineParts, List<ISoftLinePart> softLineParts, XGraphics graphics, TextMode mode, Table table)
        { 
            return lineParts.Sum(part => graphics.MeasureString(part.Text(softLineParts, mode),
                part.GetSoftLinePart(softLineParts).Span.Font(table), MeasureTrailingSpacesStringFormat).Width);
        }

        public static double GetLineCount(XGraphics graphics, Paragraph paragraph, double width, TextMode mode, Document document, Table table)
        {
            return paragraph.GetSoftLines(document, table).SelectMany(
                softLineParts => GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width),
                    GetCharInfos(softLineParts, mode), paragraph, mode, document, table)
            ).Count();
        }

        public static double GetHeight(XGraphics graphics, Paragraph paragraph, double width, TextMode mode, Document document, Table table)
        {
            return paragraph.GetSoftLines(document, table).Sum(softLineParts => {
                var charInfos = GetCharInfos(softLineParts, mode);
                return GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width), charInfos, paragraph, mode, document, table)
                    .Sum(line => paragraph.LineSpacingFunc()(line.GetLineParts(charInfos).Spans(softLineParts)
                        .Max(span => span.FontWithoutInlineVerticalAlign(table).GetHeight())));
            });
        }

        internal static IEnumerable<Span> Spans(this IEnumerable<LinePart> lineParts, List<ISoftLinePart> softLineParts)
        {
            return lineParts.Any() 
                ? lineParts.Select(part => part.GetSoftLinePart(softLineParts).Span) 
                : softLineParts.Select(softLinePart => softLinePart.Span);
        }

        internal static List<CharInfo> GetCharInfos(List<ISoftLinePart> softLineParts, TextMode mode)
        {
            return softLineParts.SelectMany(
                (part, partIndex) => part.Text(mode)
                    .Select((c, charIndex) => new CharInfo(partIndex, charIndex))).ToList();
        }

        private static double BaseLine(XFont font)
        {
            var lineSpace = font.GetHeight();
            return (lineSpace +
                    lineSpace * (font.FontFamily.GetCellAscent(font.Style) -
                        font.FontFamily.GetCellDescent(font.Style))
                    / font.FontFamily.GetLineSpacing(font.Style)) /
                2;
        }

        internal static IEnumerable<LineInfo> GetLines(XGraphics graphics, List<ISoftLinePart> softLineParts, double width, List<CharInfo> charInfos,
		    Paragraph paragraph, TextMode mode, Document document, Option<Table> table)
        {
            var runningWidths = GetRunningWidths(softLineParts, graphics, mode, table);
            var startIndex = 0;
            double previousLineWidth = 0;
            var lineIndex = 0;
            double GetWidth() => lineIndex == 0 ? width - paragraph.TextIndent().ToOption().ValueOr(0) : width;
            while (true)
            {
                var binarySearch = BinarySearch(startIndex, charInfos.Count - startIndex, i => {
                    var previousSpansWidth = charInfos[i].PartIndex == charInfos[startIndex].PartIndex
                        ? 0
                        : runningWidths[softLineParts[charInfos[i].PartIndex - 1]] - previousLineWidth;
                    if (previousSpansWidth > GetWidth()) return 1;
                    var part = softLineParts[charInfos[i].PartIndex];
                    var spanStartIndex = charInfos[i].PartIndex == charInfos[startIndex].PartIndex 
                        ? charInfos[startIndex].CharIndex 
                        : 0;
                    var text = part.Text(mode).Substring(spanStartIndex, charInfos[i].CharIndex - spanStartIndex + 1);
                    var endWidth = graphics.MeasureString(text, part.Span.Font(table), MeasureTrailingSpacesStringFormat).Width;
                    return (previousSpansWidth + endWidth).CompareTo(GetWidth());
                });
                int endIndex;
                if (binarySearch < 0)
                    if (~binarySearch == charInfos.Count)
                    {
                        yield return new LineInfo(startIndex, TrimEnd(charInfos.Count - 1, charInfos, softLineParts, startIndex, mode));
                        lineIndex++;
                        yield break;
                    }
                    else if (~binarySearch == startIndex)
                        endIndex = startIndex;
                    else
                        endIndex = ~binarySearch - 1;
                else
                    endIndex = binarySearch;
                int endIndex2;
                if (softLineParts[charInfos[endIndex].PartIndex].Span.Text(table).ExpressionVisible(document))
                    endIndex2 = endIndex + softLineParts[charInfos[endIndex].PartIndex].Text(mode).Length -
                        charInfos[endIndex].CharIndex - 1;
                else
                    endIndex2 = endIndex;
                if (endIndex2 == charInfos.Count - 1)
                {
                    yield return new LineInfo(startIndex, TrimEnd(endIndex2, charInfos, softLineParts, startIndex, mode));
                    lineIndex++;
                    yield break;
                }
                var shiftedEndIndex = ShiftEndIndex(endIndex2, charInfos, softLineParts, startIndex, mode);
                yield return new LineInfo(startIndex, TrimEnd(shiftedEndIndex, charInfos, softLineParts, startIndex, mode));
	            if (shiftedEndIndex == charInfos.Count - 1)
		            yield break;
				lineIndex++;
                startIndex = shiftedEndIndex + 1;
                var endPart = softLineParts[charInfos[shiftedEndIndex].PartIndex];
                previousLineWidth = runningWidths[endPart] -
                    graphics.MeasureString(endPart.Text(mode).Substring(charInfos[shiftedEndIndex].CharIndex + 1),
                        endPart.Span.Font(table), MeasureTrailingSpacesStringFormat).Width;
            }
        }

	    private static bool ExpressionVisible(this IText text, Document document) => document.ExpressionVisible && text.IsExpression;

	    private static int TrimEnd(int endIndex, List<CharInfo> chars, List<ISoftLinePart> softLineParts, int startIndex, TextMode mode)
        {
            var i = endIndex;
            while (i >= startIndex && IsLineBreakChar(chars.Char(i, softLineParts, mode)))
                i--;
            return i;
        }

        private static bool IsLineBreakChar(char c) => char.IsWhiteSpace(c) && c != '\u00A0';

        private static int ShiftEndIndex(int endIndex, List<CharInfo> chars, List<ISoftLinePart> softLineParts, int startIndex,
            TextMode mode)
        {
            if (IsLineBreakChar(chars.Char(endIndex + 1, softLineParts, mode)))
            {
                var i = endIndex;
                while (i < chars.Count - 1 && IsLineBreakChar(chars.Char(i + 1, softLineParts, mode)))
                    i++;
                return i;
            }
            else
            {
                int? whiteSpaceIndex = null;
                var i = endIndex;
                while (i >= startIndex)
                {
                    if (IsLineBreakChar(chars.Char(i, softLineParts, mode)))
                    {
                        whiteSpaceIndex = i;
                        break;
                    }
                    i--;
                }
                return whiteSpaceIndex ?? endIndex;
            }
        }

        private static char Char(this List<CharInfo> chars, int index, List<ISoftLinePart> softLineParts, TextMode mode)
        {
            var charInfo = chars[index];
            return softLineParts[charInfo.PartIndex].Text(mode)[charInfo.CharIndex];
        }

        private static Dictionary<ISoftLinePart, double> GetRunningWidths(List<ISoftLinePart> parts, XGraphics graphics, TextMode mode, Option<Table> table)
        {
            var runningWidth = 0d;
            return parts.Select(part => {
                runningWidth += graphics.MeasureString(part.Text(mode), part.Span.Font(table), MeasureTrailingSpacesStringFormat).Width;
                return new {part, runningWidth};
            }).ToDictionary(_ => _.part, _ => _.runningWidth);
        }

        internal class LineInfo
        {
            private int StartIndex { get; }
            private int EndIndex { get; }

            public LineInfo(int startIndex, int endIndex)
            {
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            public IEnumerable<LinePart> GetLineParts(List<CharInfo> charInfos)
            {
                if (charInfos.Count <= StartIndex) yield break;
                var info = charInfos[StartIndex];
                for (var index = StartIndex; index <= EndIndex; index++)
                {
                    var charInfo = charInfos[index];
                    if (info.PartIndex != charInfo.PartIndex)
                    {
                        yield return new LinePart(
                            info.PartIndex, info.CharIndex, charInfos[index - 1].CharIndex);
                        info = charInfo;
                    }
                }
	            if (StartIndex <= EndIndex)
		            yield return new LinePart(
			            info.PartIndex, info.CharIndex, charInfos[EndIndex].CharIndex);
            }
        }

        internal class LinePart
        {
            private int PartIndex { get; }
            private int StartIndex { get; }
            private int EndIndex { get; }
            private int Length => EndIndex - StartIndex + 1;

            public LinePart(int partIndex, int startIndex, int endIndex)
            {
                PartIndex = partIndex;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            public ISoftLinePart GetSoftLinePart(List<ISoftLinePart> softLineParts) => softLineParts[PartIndex];

            public string Text(List<ISoftLinePart> softLineParts, TextMode mode)
                => GetSoftLinePart(softLineParts).Text(mode).Substring(StartIndex, Length);

            public IText SubText(List<ISoftLinePart> softLineParts) 
                => GetSoftLinePart(softLineParts).SubText(StartIndex, Length);
        }

        internal class CharInfo
        {
            public int PartIndex { get; }
            public int CharIndex { get; }

            public CharInfo(int partIndex, int charIndex)
            {
                PartIndex = partIndex;
                CharIndex = charIndex;
            }
        }

        static ParagraphRenderer()
        {
            var xStringFormat = XStringFormats.Default;
            MeasureTrailingSpacesStringFormat =  xStringFormat;
        }

        public static XStringFormat MeasureTrailingSpacesStringFormat;

        public static void Add<TKey, TValue>(this Dictionary<TKey, List<TValue>> it, TKey key, TValue value)
        {
            if (it.TryGetValue(key, out var list))
                list.Add(value);
            else
                it.Add(key, new List<TValue> {value});
        }

        /// <summary>
        /// https://referencesource.microsoft.com/#mscorlib/system/collections/generic/arraysorthelper.cs,188
        /// </summary>
        private static int BinarySearch(int index, int length, Func<int, int> comparer)
        {
            int lo = index;
            int hi = index + length - 1;
            while (lo <= hi)
            {
                int i = lo + ((hi - lo) >> 1);
                int order = comparer(i);
                if (order == 0)
                    return i;
                if (order < 0)
                    lo = i + 1;
                else
                    hi = i - 1;
            } 
            return ~lo;
        }

        public static IEnumerable<string> SplitToLines(this string text)
        {
            return text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        }
    }
}