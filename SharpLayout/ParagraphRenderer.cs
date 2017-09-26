using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public static class ParagraphRenderer
    {
        public static void Draw(XGraphics graphics, Paragraph paragraph, XUnit x0, XUnit y0, double width, HorizontalAlign alignment,
            Drawer drawer)
        {
            var y = y0 + paragraph.TopMargin.ValueOr(0);
            foreach (var softLineParts in GetSoftLines(paragraph))
            {
                var charInfos = GetCharInfos(softLineParts);
                var innerWidth = paragraph.GetInnerWidth(width);
                var lineInfos = GetLines(graphics, softLineParts, innerWidth, charInfos).ToList();
                foreach (var line in lineInfos)
                {
                    var lineParts = line.GetLineParts(charInfos).ToList();
                    double dx;
                    switch (alignment)
                    {
                        case HorizontalAlign.Left:
                            dx = 0;
                            break;
                        case HorizontalAlign.Center:
                            dx = (innerWidth - lineParts.ContentWidth(softLineParts, graphics)) / 2;
                            break;
                        case HorizontalAlign.Right:
                            dx = innerWidth - lineParts.ContentWidth(softLineParts, graphics);
                            break;
                        default:
                            throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
                    }
                    var baseLine = lineParts.Spans(softLineParts).Max(span => BaseLine(span, graphics));
                    var x = x0 + paragraph.LeftMargin.ValueOr(0) + dx;
                    foreach (var part in lineParts)
                    {
                        var text = part.Text(softLineParts);
                        var span = part.GetSoftLinePart(softLineParts).Span;
                        drawer.DrawString(text, span.Font, span.Brush, x, y + baseLine);
                        x += graphics.MeasureString(text, span.Font, MeasureTrailingSpacesStringFormat).Width;
                    }
                    y += lineParts.Spans(softLineParts).Max(span => span.Font.LineSpace(graphics));
                }
            }
        }

        public static double GetInnerWidth(this Paragraph paragraph, double width) 
            => width - paragraph.LeftMargin.ValueOr(0) - paragraph.RightMargin.ValueOr(0);

        private static List<List<SoftLinePart>> GetSoftLines(this Paragraph paragraph)
        {
            var result = new List<List<SoftLinePart>>();
            foreach (var span in paragraph.Spans)
            {
                var lines = span.TextOrEmpty.SplitToLines().ToList();
                if (result.Count == 0)
                    result.Add(new List<SoftLinePart>());
                if (lines.Count == 0)
                    result[result.Count - 1].Add(new SoftLinePart(span, span.TextOrEmpty));
                else
                {
                    result[result.Count - 1].Add(new SoftLinePart(span, lines[0]));
                    for (var i = 1; i < lines.Count; i++)
                        result.Add(new List<SoftLinePart> {new SoftLinePart(span, lines[i])});
                }
            }
            return result;
        }

        private class SoftLinePart
        {
            public Span Span { get; }
            public string Text { get; }

            public SoftLinePart(Span span, string text)
            {
                Span = span;
                Text = text;
            }
        }

        private static double ContentWidth(this List<LinePart> lineParts, List<SoftLinePart> softLineParts, XGraphics graphics)
        { 
            return lineParts.Sum(part => graphics.MeasureString(part.Text(softLineParts),
                part.GetSoftLinePart(softLineParts).Span.Font, MeasureTrailingSpacesStringFormat).Width);
        }

        public static double GetHeight(XGraphics graphics, Paragraph paragraph, double width)
        {
            return GetSoftLines(paragraph).Sum(softLineParts => {
                var charInfos = GetCharInfos(softLineParts);
                return GetLines(graphics, softLineParts, paragraph.GetInnerWidth(width), charInfos)
                    .Sum(line => line.GetLineParts(charInfos).Spans(softLineParts)
                        .Max(span => span.Font.LineSpace(graphics)));
            });
        }

        private static IEnumerable<Span> Spans(this IEnumerable<LinePart> lineParts, List<SoftLinePart> softLineParts)
        {
            return lineParts.Any() 
                ? lineParts.Select(part => part.GetSoftLinePart(softLineParts).Span) 
                : softLineParts.Select(softLinePart => softLinePart.Span);
        }

        private static List<CharInfo> GetCharInfos(List<SoftLinePart> softLineParts)
        {
            return softLineParts.SelectMany(
                (part, partIndex) => part.Text
                    .Select((c, charIndex) => new CharInfo(partIndex, charIndex))).ToList();
        }

        private static double BaseLine(Span span, XGraphics graphics)
        {
            var lineSpace = span.Font.LineSpace(graphics);
            return (lineSpace +
                    lineSpace * (span.Font.FontFamily.GetCellAscent(span.Font.Style) -
                        span.Font.FontFamily.GetCellDescent(span.Font.Style))
                    / span.Font.FontFamily.GetLineSpacing(span.Font.Style)) /
                2;
        }

        private static IEnumerable<LineInfo> GetLines(XGraphics graphics, List<SoftLinePart> softLineParts, double width, List<CharInfo> charInfos)
        {
            var runningWidths = GetRunningWidths(softLineParts, graphics);
            var startIndex = 0;
            double previousLineWidth = 0;
            while (true)
            {
                var binarySearch = BinarySearch(startIndex, charInfos.Count - startIndex, i => {
                    var previousSpansWidth = charInfos[i].PartIndex == charInfos[startIndex].PartIndex
                        ? 0
                        : runningWidths[softLineParts[charInfos[i].PartIndex - 1]] - previousLineWidth;
                    if (previousSpansWidth > width) return 1;
                    var part = softLineParts[charInfos[i].PartIndex];
                    var spanStartIndex = charInfos[i].PartIndex == charInfos[startIndex].PartIndex 
                        ? charInfos[startIndex].CharIndex 
                        : 0;
                    var text = part.Text.Substring(spanStartIndex, charInfos[i].CharIndex - spanStartIndex + 1);
                    var endWidth = graphics.MeasureString(text, part.Span.Font, MeasureTrailingSpacesStringFormat).Width;
                    return (previousSpansWidth + endWidth).CompareTo(width);
                });
                int endIndex;
                if (binarySearch < 0)
                    if (~binarySearch == charInfos.Count)
                    {
                        yield return new LineInfo(startIndex, TrimEnd(charInfos.Count - 1, charInfos, softLineParts, startIndex));
                        yield break;
                    }
                    else if (~binarySearch == startIndex)
                        endIndex = startIndex;
                    else
                        endIndex = ~binarySearch - 1;
                else
                    endIndex = binarySearch;
                if (endIndex == charInfos.Count - 1)
                {
                    yield return new LineInfo(startIndex, TrimEnd(endIndex, charInfos, softLineParts, startIndex));
                    yield break;
                }
                var shiftedEndIndex = ShiftEndIndex(endIndex, charInfos, softLineParts, startIndex);
                yield return new LineInfo(startIndex, TrimEnd(shiftedEndIndex, charInfos, softLineParts, startIndex));
                startIndex = shiftedEndIndex + 1;
                var endPart = softLineParts[charInfos[shiftedEndIndex].PartIndex];
                previousLineWidth = runningWidths[endPart] -
                    graphics.MeasureString(endPart.Text.Substring(charInfos[shiftedEndIndex].CharIndex + 1),
                        endPart.Span.Font, MeasureTrailingSpacesStringFormat).Width;
            }
        }

        private static int TrimEnd(int endIndex, List<CharInfo> chars, List<SoftLinePart> softLineParts, int startIndex)
        {
            var i = endIndex;
            while (i >= startIndex && IsLineBreakChar(chars.Char(i, softLineParts)))
                i--;
            return i;
        }

        private static bool IsLineBreakChar(char c) => char.IsWhiteSpace(c) && c != '\u00A0';

        private static int ShiftEndIndex(int endIndex, List<CharInfo> chars, List<SoftLinePart> softLineParts, int startIndex)
        {
            if (IsLineBreakChar(chars.Char(endIndex + 1, softLineParts)))
            {
                var i = endIndex;
                while (i < chars.Count && IsLineBreakChar(chars.Char(i + 1, softLineParts)))
                    i++;
                return i;
            }
            else
            {
                int? whiteSpaceIndex = null;
                var i = endIndex;
                while (i >= startIndex)
                {
                    if (IsLineBreakChar(chars.Char(i, softLineParts)))
                    {
                        whiteSpaceIndex = i;
                        break;
                    }
                    i--;
                }
                return whiteSpaceIndex ?? endIndex;
            }
        }

        private static char Char(this List<CharInfo> chars, int index, List<SoftLinePart> softLineParts)
        {
            var charInfo = chars[index];
            return softLineParts[charInfo.PartIndex].Text[charInfo.CharIndex];
        }

        private static Dictionary<SoftLinePart, double> GetRunningWidths(List<SoftLinePart> parts, XGraphics graphics)
        {
            var runningWidth = 0d;
            return parts.Select(part => {
                runningWidth += graphics.MeasureString(part.Text, part.Span.Font, MeasureTrailingSpacesStringFormat).Width;
                return new {part, runningWidth};
            }).ToDictionary(_ => _.part, _ => _.runningWidth);
        }

        private class LineInfo
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
                return charInfos.Skip(StartIndex).Take(EndIndex - StartIndex + 1)
                    .GroupBy(charInfo => charInfo.PartIndex)
                    .Select(grouping => new LinePart(
                        PartIndex: grouping.Key, 
                        startIndex: grouping.First().CharIndex, 
                        endIndex: grouping.Last().CharIndex));
            }
        }

        private class LinePart
        {
            private int PartIndex { get; }
            private int StartIndex { get; }
            private int EndIndex { get; }

            public LinePart(int PartIndex, int startIndex, int endIndex)
            {
                this.PartIndex = PartIndex;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            public SoftLinePart GetSoftLinePart(List<SoftLinePart> softLineParts) => softLineParts[PartIndex];

            public string Text(List<SoftLinePart> softLineParts)
                => GetSoftLinePart(softLineParts).Text.Substring(StartIndex, EndIndex - StartIndex + 1);
        }

        private class CharInfo
        {
            public int PartIndex { get; }
            public int CharIndex { get; }

            public CharInfo(int partIndex, int charIndex)
            {
                PartIndex = partIndex;
                CharIndex = charIndex;
            }
        }

        public static XStringFormat MeasureTrailingSpacesStringFormat
        {
            get
            {
                var xStringFormat = XStringFormats.Default;
                xStringFormat.FormatFlags |= XStringFormatFlags.MeasureTrailingSpaces;
                return xStringFormat;
            }
        }

        public static double LineSpace(this XFont font, XGraphics graphics) => font.GetHeight(graphics);

        public static double GetSpaceWidth(XGraphics graphics, XFont font) 
            => graphics.MeasureString(" ", font, MeasureTrailingSpacesStringFormat).Width;

        public static void Add<TKey, TValue>(this Dictionary<TKey, List<TValue>> it, TKey key, TValue value)
        {
            List<TValue> list;
            if (it.TryGetValue(key, out list))
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