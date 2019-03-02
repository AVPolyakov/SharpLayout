using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public struct CharInfo
    {
        public int PartIndex { get; }
        public int CharIndex { get; }

        public CharInfo(int partIndex, int charIndex)
        {
            PartIndex = partIndex;
            CharIndex = charIndex;
        }
    }

    public struct LineInfo
    {
        private int StartIndex { get; }
        private int EndIndex { get; }

        public LineInfo(int startIndex, int endIndex)
        {
            StartIndex = startIndex;
            EndIndex = endIndex;
        }

        public List<LinePart> GetLineParts(CharInfo[] charInfos)
        {
            var list = new List<LinePart>(4);
            if (charInfos.Length <= StartIndex) return list;
            var info = charInfos[StartIndex];
            for (var index = StartIndex; index <= EndIndex; index++)
            {
                var charInfo = charInfos[index];
                if (info.PartIndex != charInfo.PartIndex)
                {
                    list.Add(new LinePart(
                        info.PartIndex, info.CharIndex, charInfos[index - 1].CharIndex));
                    info = charInfo;
                }
            }
            if (StartIndex <= EndIndex)
                list.Add(new LinePart(
                    info.PartIndex, info.CharIndex, charInfos[EndIndex].CharIndex));
            return list;
        }
    }

    public struct LinePart
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

        public IValue SubText(List<ISoftLinePart> softLineParts)
            => GetSoftLinePart(softLineParts).SubText(StartIndex, Length);
    }

    internal static class ParagraphRenderer
    {
	    public static void Draw(IGraphics graphics, Paragraph paragraph, XUnit x0, XUnit y0, double width, HorizontalAlign alignment, Drawer drawer,
		    GraphicsType graphicsType, TextMode mode, Document document, Table table,
	        DrawCache drawCache)
        {
            var y = y0 + paragraph.TopMargin().ToOption().ValueOr(0);
            var lineCount = Lazy.Create(() => GetLineCount(graphics, paragraph, width, mode, document, table, drawCache));
            var lineIndex = 0;
            double TextIndent() => lineIndex == 0 ? paragraph.TextIndent().ToOption().ValueOr(0) : 0;
            var softLines = paragraph.GetSoftLines(document, table, drawCache, mode);
            var leftIndent = x0 + paragraph.LeftMargin().ToOption().ValueOr(0);
            for (var softLineIndex = 0; softLineIndex < softLines.Count; softLineIndex++)
            {
                var softLineParts = softLines[softLineIndex];
                var charInfos = GetCharInfos(softLineParts, mode);
                var innerWidth = paragraph.GetInnerWidth(width);
                foreach (var line in GetLines(graphics, softLineParts, innerWidth, charInfos, paragraph, mode, document, table))
                {
                    var lineParts = line.GetLineParts(charInfos);
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
                    var spans = lineParts.Spans(softLineParts);
                    var baseLine = spans.Max(span => BaseLine(span.FontWithoutInlineVerticalAlign(table), drawCache));
                    var x = leftIndent + TextIndent() + dx;
                    var maxLineSpace = spans.Max(span => span.FontWithoutInlineVerticalAlign(table).GetHeight());
                    var multiplier = Lazy.Create(() => {
                        var spaces = GetSpaces(lineParts, softLineParts, mode);
                        return spaces.Any()
                            ? (innerWidth - lineParts.ContentWidth(softLineParts, graphics, mode, table) - TextIndent()) /
                              spaces.Sum(tuple => graphics.MeasureString(new string(tuple.Item2.Char, 1),
                                  tuple.Item1.GetSoftLinePart(softLineParts).Span.Font(table).XFont, MeasureTrailingSpacesStringFormat).Width)
                            : new double?();
                    });
                    foreach (var part in lineParts)
                    {
                        var text = part.Text(softLineParts, mode);
                        var span = part.GetSoftLinePart(softLineParts).Span;
                        var rectangleWidth = 0d;
                        var rectangleX = x;
                        Font font;
                        if (alignment == HorizontalAlign.Justify && span.Font(table).Underline)
                        {
                            font = new Font(span.Font(table).FontFamily.Name, span.Font(table).Size, new[] {
                                XFontStyle.Regular,
                                XFontStyle.Bold,
                                XFontStyle.Italic,
                                XFontStyle.Strikeout,
                            }.Where(_ => span.Font(table).Style.HasFlag(_)).Aggregate((style1, style2) => style1 | style2), span.Font(table).PdfOptions);
                        }
                        else
                            font = span.Font(table);
                        if (alignment == HorizontalAlign.Justify || graphicsType == GraphicsType.Image && !font.Underline)
                            foreach (var drawTextPart in GetDrawTextParts(text))
                            {
                                double stringWidth;
                                switch (drawTextPart)
                                {
                                    case DrawTextPart.Space space:
                                        var measureWidth = graphics.MeasureString(new string(space.Char, 1), font.XFont, MeasureTrailingSpacesStringFormat).Width;
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
                                        drawer.DrawString(word.Text, font, span.CalculateBrush(document, table), x, y + baseLine + span.InlineVerticalAlignOffset(table, drawCache));
                                        stringWidth = graphics.MeasureString(word.Text, font.XFont, MeasureTrailingSpacesStringFormat).Width;
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }
                                x += stringWidth;
                                rectangleWidth += stringWidth;
                            }
                        else
                        {
                            var measureString = graphics.MeasureString(text, font.XFont, MeasureTrailingSpacesStringFormat);
                            drawer.DrawString(text, font, span.CalculateBrush(document, table), x, y + baseLine + span.InlineVerticalAlignOffset(table, drawCache));
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
                                drawer.DrawLine(new XPen(solidBrush.Color, font.Size * widthMultiplier),
                                    rectangleX, lineY, rectangleX + rectangleWidth, lineY);
                            }
                    }
                    y += paragraph.LineSpacingFunc()(maxLineSpace);
                    lineIndex++;
                }
            }
        }

        private static double InlineVerticalAlignOffset(this Span span, Table table, DrawCache drawCache)
        {
            switch (span.InlineVerticalAlign())
            {
                case InlineVerticalAlign.Baseline:
                    return 0;
                case InlineVerticalAlign.Sub:
                    return -BaseLine(span.FontWithoutInlineVerticalAlign(table), drawCache) + BaseLine(span.Font(table), drawCache) +
                        span.FontWithoutInlineVerticalAlign(table).GetHeight() - span.Font(table).GetHeight();
                case InlineVerticalAlign.Super:
                    return -BaseLine(span.FontWithoutInlineVerticalAlign(table), drawCache) + BaseLine(span.Font(table), drawCache);
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

        internal static List<List<ISoftLinePart>> GetSoftLines(this Paragraph paragraph, Document document, Option<Table> table, DrawCache drawCache, TextMode mode)
        {
            var result = new List<List<ISoftLinePart>>();
            foreach (var span in paragraph.Spans)
            {
                if (result.Count == 0)
                    result.Add(new List<ISoftLinePart>(4));
                var parts = span.GetSoftLineParts(span, document, drawCache, table, mode);
                result[result.Count - 1].Add(parts[0]);
                for (var i = 1; i < parts.Length; i++)
                    result.Add(new List<ISoftLinePart> {parts[i]});
            }
            return result;
        }

        private static double ContentWidth(this List<LinePart> lineParts, List<ISoftLinePart> softLineParts, IGraphics graphics, TextMode mode, Table table)
        { 
            return lineParts.Sum(part => graphics.MeasureString(part.Text(softLineParts, mode),
                part.GetSoftLinePart(softLineParts).Span.Font(table).XFont, MeasureTrailingSpacesStringFormat).Width);
        }

        public static double GetLineCount(IGraphics graphics, Paragraph paragraph, double width, TextMode mode, Document document, Table table, DrawCache drawCache)
        {
            return paragraph.GetSoftLines(document, table, drawCache, mode).Select((softLineParts, i) => (softLineParts, i)).Select(
                _ => GetLines(graphics, _.softLineParts, paragraph.GetInnerWidth(width), GetCharInfos(_.softLineParts, mode), paragraph, mode, document, table).Count
            ).Sum();
        }

        public static double GetHeight(IGraphics graphics, Paragraph paragraph, double width, TextMode mode, Document document, Table table, DrawCache drawCache)
        {
            var softLines = paragraph.GetSoftLines(document, table, drawCache, mode);
            double sum = 0;
            for (var index = 0; index < softLines.Count; index++)
            {
                var softLineParts = softLines[index];
                var charInfos = GetCharInfos(softLineParts, mode);
                foreach (var line in GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width), charInfos, paragraph, mode, document, table))
                {
                    double max = 0;
                    var spans = Spans(line.GetLineParts(charInfos), softLineParts);
                    for (var i = 0; i < spans.Length; i++)
                    {
                        var height = spans[i].FontWithoutInlineVerticalAlign(table).GetHeight();
                        if (max < height) max = height;
                    }
                    sum += paragraph.LineSpacingFunc()(max);
                }
            }
            return sum;
        }

        internal static Span[] Spans(this List<LinePart> lineParts, List<ISoftLinePart> softLineParts)
        {
            if (lineParts.Count > 0)
            {
                var spans = new Span[lineParts.Count];
                for (var index = 0; index < lineParts.Count; index++)
                    spans[index] = lineParts[index].GetSoftLinePart(softLineParts).Span;
                return spans;
            }
            else
            {
                var spans = new Span[softLineParts.Count];
                for (var index = 0; index < softLineParts.Count; index++)
                    spans[index] = softLineParts[index].Span;
                return spans;
            }
        }

        internal static CharInfo[] GetCharInfos(List<ISoftLinePart> softLineParts, TextMode mode)
        {
            var list = new CharInfo[softLineParts.Sum(_ => _.Text(mode).Length)];
            var i = 0;
            for (var partIndex = 0; partIndex < softLineParts.Count; partIndex++)
            {
                var text = softLineParts[partIndex].Text(mode);
                for (var charIndex = 0; charIndex < text.Length; charIndex++)
                    list[i++] = new CharInfo(partIndex, charIndex);
            }
            return list;
        }

        private static double BaseLine(Font font, DrawCache drawCache)
        {
            var key = new FontKey(font.Name, font.Size, font.Style, font.PdfOptions.FontEncoding);
            if (drawCache.BaseLines.TryGetValue(key, out var value)) 
                return value;
            var lineSpace = font.GetHeight();
            var baseLine = (lineSpace +
                            lineSpace * (font.FontFamily.GetCellAscent(font.Style) -
                                         font.FontFamily.GetCellDescent(font.Style))
                            / font.FontFamily.GetLineSpacing(font.Style)) /
                           2;
            drawCache.BaseLines.Add(key, baseLine);
            return baseLine;
        }

        internal static List<LineInfo> GetLines(IGraphics graphics, List<ISoftLinePart> softLineParts, double width, CharInfo[] charInfos,
		    Paragraph paragraph, TextMode mode, Document document, Option<Table> table)
        {
            var result = new List<LineInfo>(4);
            var startIndex = 0;
            var lineIndex = 0;
            double GetWidth() => lineIndex == 0 ? width - paragraph.TextIndent().ToOption().ValueOr(0) : width;
            while (true)
            {
                var endIndex = GetEndIndex(GetWidth(), charInfos, softLineParts, mode, startIndex, graphics, table);
                if (endIndex == charInfos.Length - 1)
                {
                    result.Add(new LineInfo(startIndex, TrimEnd(charInfos.Length - 1, charInfos, softLineParts, startIndex, mode)));
                    return result;
                }
                int endIndex2;
                if (softLineParts[charInfos[endIndex].PartIndex].Span.Text(table).ExpressionVisible(document))
                    endIndex2 = endIndex + softLineParts[charInfos[endIndex].PartIndex].Text(mode).Length -
                        charInfos[endIndex].CharIndex - 1;
                else
                    endIndex2 = endIndex;
                if (endIndex2 == charInfos.Length - 1)
                {
                    result.Add(new LineInfo(startIndex, TrimEnd(endIndex2, charInfos, softLineParts, startIndex, mode)));
                    return result;
                }
                var shiftedEndIndex = ShiftEndIndex(endIndex2, charInfos, softLineParts, startIndex, mode);
                result.Add(new LineInfo(startIndex, TrimEnd(shiftedEndIndex, charInfos, softLineParts, startIndex, mode)));
                if (shiftedEndIndex == charInfos.Length - 1)
                    return result;
				lineIndex++;
                startIndex = shiftedEndIndex + 1;
            }
        }

        private static int GetEndIndex(double width, CharInfo[] charInfos, List<ISoftLinePart> softLineParts, TextMode mode, int startIndex, IGraphics graphics, Option<Table> table)
        {
            var i = startIndex;
            double currentWidth = 0;
            while (true)
            {
                if (i >= charInfos.Length) return charInfos.Length - 1;
                var charInfo = charInfos[i];
                var part = softLineParts[charInfo.PartIndex];
                currentWidth += part.Text(mode).GetCharWidth(charInfo.CharIndex, part, graphics, table);
                if (currentWidth >= width)
                    if (i == startIndex)
                        return i;
                    else
                        return i - 1;
                i++;
            }
        }

        private static double GetCharWidth(this string text, int charIndex, ISoftLinePart part, IGraphics graphics, Option<Table> table)
        {
            int key = text[charIndex];
            if (part.CharSizeCache.Dictionary.TryGetValue(key, out var value)) 
                return value;
            var width = graphics.MeasureString(text.Substring(charIndex, 1), part.Span.Font(table).XFont, MeasureTrailingSpacesStringFormat).Width;
            part.CharSizeCache.Dictionary.Add(key, width);
            return width;
        }

        private static bool ExpressionVisible(this IValue value, Document document) => document.ExpressionVisible && value.IsExpression;

	    private static int TrimEnd(int endIndex, CharInfo[] chars, List<ISoftLinePart> softLineParts, int startIndex, TextMode mode)
        {
            var i = endIndex;
            while (i >= startIndex && IsLineBreakChar(chars.Char(i, softLineParts, mode)))
                i--;
            return i;
        }

        private static bool IsLineBreakChar(char c) => char.IsWhiteSpace(c) && c != '\u00A0';

        private static int ShiftEndIndex(int endIndex, CharInfo[] chars, List<ISoftLinePart> softLineParts, int startIndex,
            TextMode mode)
        {
            if (IsLineBreakChar(chars.Char(endIndex + 1, softLineParts, mode)))
            {
                var i = endIndex;
                while (i < chars.Length - 1 && IsLineBreakChar(chars.Char(i + 1, softLineParts, mode)))
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

        private static char Char(this CharInfo[] chars, int index, List<ISoftLinePart> softLineParts, TextMode mode)
        {
            var charInfo = chars[index];
            return softLineParts[charInfo.PartIndex].Text(mode)[charInfo.CharIndex];
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

        public static string[] SplitToLines(this string text)
        {
            return text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        }
    }
}