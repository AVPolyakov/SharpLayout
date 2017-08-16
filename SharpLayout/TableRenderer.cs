using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public static class TableRenderer
    {
        public static int Draw(XGraphics xGraphics, PageSettings pageSettings, Action<int, Action<XGraphics>> pageAction,
            IEnumerable<Table> tables)
        {
            var firstOnPage = true;
            var y = pageSettings.TopMargin;
            var tableParts = tables.Select(table => {
                var tableInfo = GetTableInfo(xGraphics, table, y);
                double endY;
                var rowSets = SplitByPages(tableInfo, firstOnPage, out endY, pageSettings);
                if (rowSets.Count > 0)
                    firstOnPage = false;
                y = endY;
                return rowSets.Select((rows, index) => new TablePart(rows, index, tableInfo));
            }).ToList();
            if (tableParts.Count == 0) return 0;
            var pages = new List<List<TablePart>>();
            foreach (var part in tableParts.SelectMany(_ => _))
                if (pages.Count > 0)
                    if (part.IsFirst)
                        pages[pages.Count - 1].Add(part);
                    else
                        pages.Add(new List<TablePart> {part});
                else
                    pages.Add(new List<TablePart> {part});
            for (var index = 0; index < pages.Count; index++)
                if (index == 0)
                    foreach (var part in pages[index])
                        Draw(part.TableInfo, part.Rows, part.Y(pageSettings), xGraphics, pageSettings);
                else
                    pageAction(index, xGraphics2 => {
                        foreach (var part in pages[index])
                            Draw(part.TableInfo, part.Rows, part.Y(pageSettings), xGraphics2, pageSettings);
                    });
            return pages.Count;
        }

        private static List<IEnumerable<int>> SplitByPages(TableInfo tableInfo, bool firstOnPage, out double endY, PageSettings pageSettings)
        {
            if (tableInfo.Table.Rows.Count == 0)
            {
                endY = tableInfo.Y;
                return new List<IEnumerable<int>>();
            }
            var mergedRows = MergedRows(tableInfo.Table);
            var y = tableInfo.Y + tableInfo.Table.Columns.Max(column => tableInfo.TopBorderFunc(new CellInfo(0, column.Index)).ValueOr(0));
            var lastRowOnPreviousPage = new Option<int>();
            var row = 0;
            var tableFirstPage = true;
            var result = new List<IEnumerable<int>>();
            while (true)
            {
                y += tableInfo.MaxHeights[row];
                if (pageSettings.PageHeight - pageSettings.BottomMargin - y < 0)
                {
                    var firstMergedRow = FirstMergedRow(mergedRows, row);
                    var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
                    if (firstMergedRow - start > 0)
                    {
                        result.Add(Enumerable.Range(start, firstMergedRow - start));
                        lastRowOnPreviousPage = firstMergedRow - 1;
                        row = firstMergedRow;
                    }
                    else
                    {
                        if (firstMergedRow == 0 && tableFirstPage && !firstOnPage)
                        {
                            result.Add(Enumerable.Empty<int>());
                            lastRowOnPreviousPage = new Option<int>();
                            row = 0;
                        }
                        else
                        {
                            var endMergedRow = EndMergedRow(tableInfo.Table, mergedRows, row);
                            result.Add(Enumerable.Range(start, endMergedRow - start));
                            lastRowOnPreviousPage = endMergedRow;
                            row = endMergedRow + 1;
                            if (row >= tableInfo.Table.Rows.Count) break;
                        }
                    }
                    tableFirstPage = false;
                    y = pageSettings.TopMargin + (row == 0
                        ? tableInfo.Table.Columns.Max(column => tableInfo.TopBorderFunc(new CellInfo(row, column.Index)).ValueOr(0))
                        : tableInfo.Table.Columns.Max(column => tableInfo.BottomBorderFunc(new CellInfo(row - 1, column.Index)).ValueOr(0)));
                }
                else
                {
                    row++;
                    if (row >= tableInfo.Table.Rows.Count) break;
                }
            }
            {
                var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
                if (start < tableInfo.Table.Rows.Count)
                    result.Add(Enumerable.Range(start, tableInfo.Table.Rows.Count - start));
            }
            endY = y;
            return result;
        }

        private static void Draw(TableInfo info, IEnumerable<int> rows, double y0, XGraphics xGraphics, PageSettings pageSettings)
        {
            var firstRow = rows.FirstOrNone();
            if (!firstRow.HasValue) return;
            var maxTopBorder = firstRow.Value == 0
                ? info.Table.Columns.Max(column => info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index)).ValueOr(0))
                : MaxTopBorder(firstRow.Value - 1, info.Table, info.BottomBorderFunc);
            {
                var x = info.Table.X0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var topBorder = firstRow.Value == 0
                        ? info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index))
                        : info.BottomBorderFunc(new CellInfo(firstRow.Value - 1, column.Index));
                    if (topBorder.HasValue)
                    {
                        var borderY = y0 + maxTopBorder - topBorder.Value/2;
                        var leftBorder = column.Index == 0
                            ? info.LeftBorderFunc(new CellInfo(firstRow.Value, 0)).ValueOr(0)
                            : info.RightBorderFunc(new CellInfo(firstRow.Value, column.Index - 1)).ValueOr(0);
                        xGraphics.DrawLine(new XPen(XColors.Black, topBorder.Value),
                            x - leftBorder, borderY,
                            x + column.Width, borderY);
                    }
                    x += column.Width;
                }
            }
            var y = y0 + maxTopBorder;
            foreach (var row in rows)
            {
                var leftBorder = info.LeftBorderFunc(new CellInfo(row, 0));
                if (leftBorder.HasValue)
                {
                    var borderX = info.Table.X0 + info.MaxLeftBorder - leftBorder.Value/2;
                    xGraphics.DrawLine(new XPen(XColors.Black, leftBorder.Value),
                        borderX, y, borderX, y + info.MaxHeights[row]);
                }
                var x = info.Table.X0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var bottomBorder = info.BottomBorderFunc(new CellInfo(row, column.Index));
                    if (pageSettings.IsHighlightCells)
                        HighlightCells(xGraphics, info, bottomBorder, row, column, x, y);
                    var cell = info.Table.Rows[row][column];
                    if (cell.Paragraph.HasValue)
                    {
                        var width = info.Table.ContentWidth(row, column, info.RightBorderFunc);
                        var rowspan = cell.Rowspan.ValueOr(1);
                        var cellInnerHeight = Enumerable.Range(0, rowspan).Sum(i => info.MaxHeights[row + i]
                            - MaxTopBorder(row + rowspan - 1, info.Table, info.BottomBorderFunc));
                        var paragraphHeight = cell.Paragraph.Value.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc);
                        double dy;
                        switch (cell.VerticalAlignment)
                        {
                            case VerticalAlignment.Top:
                                dy = 0;
                                break;
                            case VerticalAlignment.Center:
                                dy = (cellInnerHeight - paragraphHeight) / 2;
                                break;
                            case VerticalAlignment.Bottom:
                                dy = cellInnerHeight - paragraphHeight;
                                break;
                            default:
                                throw new ArgumentOutOfRangeException();
                        }
                        if (pageSettings.IsHighlightCells)
                            HighlightParagraph(cell.Paragraph.Value, column, row, x, y + dy, width, info, xGraphics);
                        ParagraphRenderer.Draw(xGraphics, cell.Paragraph.Value, x, y + dy, width, cell.Paragraph.Value.Alignment);
                    }
                    var rightBorder = info.RightBorderFunc(new CellInfo(row, column.Index));
                    if (rightBorder.HasValue)
                    {
                        var borderX = x + column.Width - rightBorder.Value/2;
                        xGraphics.DrawLine(new XPen(XColors.Black, rightBorder.Value),
                            borderX, y, borderX, y + info.MaxHeights[row]);
                    }
                    if (bottomBorder.HasValue)
                    {
                        var borderY = y + info.MaxHeights[row] - bottomBorder.Value/2;
                        xGraphics.DrawLine(new XPen(XColors.Black, bottomBorder.Value),
                            x, borderY, x + column.Width, borderY);
                    }
                    x += column.Width;
                }
                y += info.MaxHeights[row];
            }
        }

        private static double MaxTopBorder(int rowIndex, Table infoTable, Func<CellInfo, Option<double>> bottomBorderFunc)
        {
            return infoTable.Columns.Max(column => bottomBorderFunc(new CellInfo(rowIndex, column.Index)).ValueOr(0));
        }

        private static int EndMergedRow(Table table, HashSet<int> mergedRows, int row)
        {
            if (row + 1 >= table.Rows.Count) return row;
            var i = row + 1;
            while (true)
            {
                if (!mergedRows.Contains(i))
                    return i - 1;
                i++;
            }
        }

        private static int FirstMergedRow(HashSet<int> mergedRows, int row)
        {
            var i = row;
            while (true)
            {
                if (!mergedRows.Contains(i))
                    return i;
                i--;
            }
        }

        private static HashSet<int> MergedRows(Table table)
        {
            var set = new HashSet<int>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan);
                    if (rowspan.HasValue)
                        for (var i = row.Index + 1; i < row.Index + rowspan.Value; i++)
                            set.Add(i);
                }
            return set;
        }

        private static double ContentWidth(this Table table, int row, Column column, Func<CellInfo, Option<double>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Colspan).Match(
                colspan => column.Width
                    + Enumerable.Range(column.Index + 1, colspan - 1).Sum(i => table.Columns[i].Width)
                    - table.BorderWidth(row, column, column.Index + colspan - 1, rightBorderFunc),
                () => column.Width - table.BorderWidth(row, column, column.Index, rightBorderFunc));

        private static double BorderWidth(this Table table, int row, Column column, int rightColumn, Func<CellInfo, Option<double>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Rowspan).Match(
                rowspan => Enumerable.Range(row, rowspan)
                    .Max(i => rightBorderFunc(new CellInfo(i, rightColumn)).ValueOr(0)),
                () => rightBorderFunc(new CellInfo(row, rightColumn)).ValueOr(0));

        private static Dictionary<int, double> MaxHeights(this Table table, XGraphics graphics, Func<CellInfo, Option<double>> rightBorderFunc,
            Func<CellInfo, Option<double>> bottomBorderFunc)
        {
            var cellContentsByBottomRow = new Dictionary<CellInfo, MaxHeightsTuple>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var paragraph = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Paragraph);
                    if (paragraph.HasValue)
                    {
                        var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan);
                        var rowIndex = rowspan.Match(value => row.Index + value - 1, () => row.Index);
                        cellContentsByBottomRow.Add(new CellInfo(rowIndex, column.Index),
                            new MaxHeightsTuple(paragraph.Value, rowspan, row));
                    }
                }
            var result = new Dictionary<int, double>();
            foreach (var row in table.Rows)
            {
                var maxHeight = 0d;
                foreach (var column in table.Columns)
                {
                    double rowHeightByContent;
                    MaxHeightsTuple cell;
                    if (cellContentsByBottomRow.TryGetValue(new CellInfo(row, column), out cell))
                    {
                        var paragraphHeight = cell.Paragraph.GetParagraphHeight(cell.Row.Index, column, table, graphics, rightBorderFunc);
                        rowHeightByContent = cell.Rowspan.Match(
                            _ => Math.Max(paragraphHeight - Enumerable.Range(1, _ - 1).Sum(i => result[row.Index - i]), 0),
                            () => paragraphHeight);
                    }
                    else
                        rowHeightByContent = 0;
                    var innerHeight = row.Height.Match(
                        _ => Math.Max(rowHeightByContent, _), () => rowHeightByContent);
                    var height = innerHeight + MaxTopBorder(row.Index, table, bottomBorderFunc);
                    if (maxHeight < height)
                        maxHeight = height;
                }
                result.Add(row.Index, maxHeight);
            }
            return result;
        }

        private static double GetParagraphHeight(this Paragraph paragraph, int row, Column column, Table table, XGraphics graphics,
            Func<CellInfo, Option<double>> rightBorderFunc)
        {
            return paragraph.GetInnerHeight(graphics, table, row, column, rightBorderFunc) +
                paragraph.TopMargin.ValueOr(0) + paragraph.BottomMargin.ValueOr(0);
        }

        private static double GetInnerHeight(this Paragraph paragraph, XGraphics graphics, Table table, int row, Column column,
            Func<CellInfo, Option<double>> rightBorderFunc)
        {
            return ParagraphRenderer.GetHeight(graphics, paragraph, table.ContentWidth(row, column, rightBorderFunc));
        }

        private static Func<CellInfo, Option<double>> RightBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var rightBorder = table.Find(new CellInfo(row, column)).SelectMany(_ => _.RightBorder);
                    if (rightBorder.HasValue)
                    {
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                        var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan);
                        if (rowspan.HasValue)
                            for (var i = 1; i <= rowspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                                    new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                    }
                    var leftBorder = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.LeftBorder);
                    if (leftBorder.HasValue)
                    {
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
                        var rowspan = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.Rowspan);
                        if (rowspan.HasValue)
                            for (var i = 1; i <= rowspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                                    new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
                    }
                }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The right border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToSttring(table)}");
                else
                    return list[0].Width;
            });
        }

        private static Func<CellInfo, Option<double>> BottomBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var bottomBorder = row[column].BottomBorder;
                    if (bottomBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index), 
                            new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                        var colspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan);
                        if (colspan.HasValue)
                            for (var i = 1; i <= colspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                                    new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                    }
                    var topBorder = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.TopBorder);
                    if (topBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index),
                            new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
                        var colspan = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.Colspan);
                        if (colspan.HasValue)
                            for (var i = 1; i <= colspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                                    new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
                    }
                }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The bottom border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToSttring(table)}");
                else
                    return list[0].Width;
            });
        }

        private static Func<CellInfo, Option<double>> LeftBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
            {
                var leftBorder = table.Find(new CellInfo(row.Index, 0)).SelectMany(_ => _.LeftBorder);
                if (leftBorder.HasValue)
                {
                    result.Add(new CellInfo(row.Index, 0), new BorderTuple(leftBorder.Value, new CellInfo(row.Index, 0)));
                    var rowspan = table.Find(new CellInfo(row.Index, 0)).SelectMany(_ => _.Rowspan);
                    if (rowspan.HasValue)
                        for (var i = 1; i <= rowspan.Value - 1; i++)
                            result.Add(new CellInfo(row.Index + i, 0),
                                new BorderTuple(leftBorder.Value, new CellInfo(row.Index, 0)));
                }
            }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The left border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToSttring(table)}");
                else
                    return list[0].Width;
            });
        }

        private static Func<CellInfo, Option<double>> TopBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var column in table.Columns)
            {
                var bottomBorder = table.Find(new CellInfo(0, column.Index)).SelectMany(_ => _.TopBorder);
                if (bottomBorder.HasValue)
                {
                    result.Add(new CellInfo(0, column.Index),
                        new BorderTuple(bottomBorder.Value, new CellInfo(0, column.Index)));
                    var colspan = table.Find(new CellInfo(0, column.Index)).SelectMany(_ => _.Colspan);
                    if (colspan.HasValue)
                        for (var i = 1; i <= colspan.Value - 1; i++)
                            result.Add(new CellInfo(0, column.Index + i),
                                new BorderTuple(bottomBorder.Value, new CellInfo(0, column.Index)));
                }
            }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The top border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToSttring(table)}");
                else
                    return list[0].Width;
            });
        }

        private static string CellsToSttring(this IEnumerable<CellInfo> cells, Table table) 
            => string.Join(", ", cells.Select(_ => $"L{table.Line} r{_.RowIndex + 1}c{_.ColumnIndex + 1}"));

        private static TableInfo GetTableInfo(XGraphics xGraphics, Table table, double y)
        {
            var rightBorderFunc = table.RightBorder();
            var bottomBorderFunc = table.BottomBorder();
            return new TableInfo(table, table.TopBorder(), bottomBorderFunc,
                table.MaxHeights(xGraphics, rightBorderFunc, bottomBorderFunc), y, table.LeftBorder(),
                rightBorderFunc);
        }

        private class TablePart
        {
            public IEnumerable<int> Rows { get; }
            private int Index { get; }
            public TableInfo TableInfo { get; }
            public bool IsFirst => Index == 0;
            public double Y(PageSettings pageSettings) => IsFirst ? TableInfo.Y : pageSettings.TopMargin;

            public TablePart(IEnumerable<int> rows, int index, TableInfo tableInfo)
            {
                Rows = rows;
                Index = index;
                TableInfo = tableInfo;
            }
        }

        private class TableInfo
        {
            public Table Table { get; }
            public Func<CellInfo, Option<double>> RightBorderFunc { get; }
            public Func<CellInfo, Option<double>> LeftBorderFunc { get; }
            public Func<CellInfo, Option<double>> TopBorderFunc { get; }
            public Func<CellInfo, Option<double>> BottomBorderFunc { get; }
            public Dictionary<int, double> MaxHeights { get; }
            public double MaxLeftBorder { get; }
            public double Y { get; }

            public TableInfo(Table table, Func<CellInfo, Option<double>> topBorderFunc, Func<CellInfo, Option<double>> bottomBorderFunc,
                Dictionary<int, double> maxHeights, double y, Func<CellInfo, Option<double>> leftBorderFunc, Func<CellInfo, Option<double>> rightBorderFunc)
            {
                Table = table;
                RightBorderFunc = rightBorderFunc;
                LeftBorderFunc = leftBorderFunc;
                TopBorderFunc = topBorderFunc;
                BottomBorderFunc = bottomBorderFunc;
                MaxHeights = maxHeights;
                MaxLeftBorder = table.Rows.Count == 0 ? 0 : table.Rows.Max(row => leftBorderFunc(new CellInfo(row.Index, 0)).ValueOr(0));
                Y = y;                
            }
        }

        private class MaxHeightsTuple
        {
            public Paragraph Paragraph { get; }
            public Option<int> Rowspan { get; }
            public Row Row { get; }

            public MaxHeightsTuple(Paragraph paragraph, Option<int> rowspan, Row row)
            {
                Paragraph = paragraph;
                Rowspan = rowspan;
                Row = row;
            }
        }

        public class BorderTuple
        {
            public double Width { get; }
            public CellInfo CellInfo { get; }

            public BorderTuple(double width, CellInfo cellInfo)
            {
                Width = width;
                CellInfo = cellInfo;
            }
        }

        private static void HighlightParagraph(Paragraph paragraph, Column column, int row, double x, double y, double width, TableInfo info, XGraphics xGraphics)
        {
            var innerHeight = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc);
            var innerWidth = paragraph.GetInnerWidth(width);
            if (innerWidth > 0 && innerHeight > 0)
                FillRectangle(xGraphics, XColor.FromArgb(32, 0, 0, 255),
                    x + paragraph.LeftMargin.ValueOr(0),
                    y + paragraph.TopMargin.ValueOr(0),
                    innerWidth, innerHeight);
        }

        private static void HighlightCells(XGraphics xGraphics, TableInfo info, Option<double> bottomBorder, int row, Column column, double x, double y)
        {
            var color = (row + column.Index) % 2 == 1
                ? XColor.FromArgb(32, 127, 127, 127)
                : XColor.FromArgb(32, 0, 255, 0);
            var height = info.MaxHeights[row] - bottomBorder.ValueOr(0);
            var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).ValueOr(0);
            FillRectangle(xGraphics, color, x, y, width, height);
            var xFont = new XFont("Times New Roman", 10, XFontStyle.Regular,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            var xSolidBrush = new XSolidBrush(XColor.FromArgb(128, 255, 0, 0));
            if (row == 0 && column.Index == 0)
                xGraphics.DrawString($"{info.Table.Line}",
                    xFont,
                    new XSolidBrush(XColor.FromArgb(128, 0, 0, 255)),
                    new XRect(x - 100 - 2 - 10, y, 100, 100),
                    new XStringFormat {
                        Alignment = XStringAlignment.Far,
                        LineAlignment = XLineAlignment.Near
                    });
            if (column.Index == 0)
                xGraphics.DrawString($"r{row + 1}",
                    xFont,
                    xSolidBrush,
                    new XRect(x - 100 - 2, y, 100, 100),
                    new XStringFormat {
                        Alignment = XStringAlignment.Far,
                        LineAlignment = XLineAlignment.Near
                    });
            if (row == 0)
                xGraphics.DrawString($"c{column.Index + 1}",
                    xFont,
                    xSolidBrush,
                    new XRect(x, y - 100, 100, 100),
                    new XStringFormat {
                        Alignment = XStringAlignment.Near,
                        LineAlignment = XLineAlignment.Far
                    });
        }

        private static void FillRectangle(XGraphics xGraphics, XColor color, double x, double y, double width, double height)
        {
            var lineY = y + height / 2;
            //DrawRectangle does not draw the last pixel sometimes
            xGraphics.DrawLine(new XPen(color, height), x, lineY, x + width, lineY);
        }
    }
}