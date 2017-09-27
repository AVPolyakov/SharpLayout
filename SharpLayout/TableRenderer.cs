using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static System.Linq.Enumerable;

namespace SharpLayout
{
    public static class TableRenderer
    {
        public static List<SyncPageInfo> Draw(XGraphics xGraphics, PageSettings pageSettings, Action<int, Action<XGraphics>> pageAction, IEnumerable<Table> tables,
            Document document)
        {
            var tableInfos = new Dictionary<Table, TableInfo>();
            var firstOnPage = true;
            var y = pageSettings.TopMargin;
            var tableParts = tables.Select(table => {
                var tableInfo = GetTableInfo(xGraphics, table, tableInfos);
                double endY;
                var tableY = y;
                var rowSets = SplitByPages(tableInfo, firstOnPage, out endY, pageSettings, tableY);
                if (rowSets.Count > 0)
                    firstOnPage = false;
                y = endY;
                return rowSets.Select((rows, index) => new TablePart(rows, index, tableInfo, tableY));
            }).ToList();
            if (tableParts.Count == 0) return new List<SyncPageInfo>();
            var pages = new List<List<TablePart>>();
            foreach (var part in tableParts.SelectMany(_ => _))
                if (pages.Count > 0)
                    if (part.IsFirst)
                        pages[pages.Count - 1].Add(part);
                    else
                        pages.Add(new List<TablePart> {part});
                else
                    pages.Add(new List<TablePart> {part});
            var syncPageInfos = new List<SyncPageInfo>();
            for (var index = 0; index < pages.Count; index++)
            {
                var syncPageInfo = new SyncPageInfo();
                syncPageInfos.Add(syncPageInfo);
                if (index == 0)
                {
                    var drawer = new Drawer(xGraphics);
                    foreach (var part in pages[index])
                        Draw(part.TableInfo, part.Rows, part.Y(pageSettings), xGraphics, document, tableInfos,
                            pageSettings.LeftMargin, syncPageInfo, 0, drawer);
                    drawer.Flush();
                }
                else
                    pageAction(index, xGraphics2 => {
                        var drawer = new Drawer(xGraphics2);
                        foreach (var part in pages[index])
                            Draw(part.TableInfo, part.Rows, part.Y(pageSettings), xGraphics2, document, tableInfos,
                                pageSettings.LeftMargin, syncPageInfo, 0, drawer);
                        drawer.Flush();
                    });
            }
            return syncPageInfos;
        }

        private static List<IEnumerable<int>> SplitByPages(TableInfo tableInfo, bool firstOnPage, out double endY, PageSettings pageSettings, double tableY)
        {
            if (tableInfo.Table.Rows.Count == 0)
            {
                endY = tableY;
                return new List<IEnumerable<int>>();
            }
            var mergedRows = MergedRows(tableInfo.Table);
            var y = tableY + tableInfo.Table.Columns.Max(column => tableInfo.TopBorderFunc(new CellInfo(0, column.Index)).ValueOr(0));
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
                        result.Add(Range(start, firstMergedRow - start));
                        lastRowOnPreviousPage = firstMergedRow - 1;
                        row = firstMergedRow;
                    }
                    else
                    {
                        if (firstMergedRow == 0 && tableFirstPage && !firstOnPage)
                        {
                            result.Add(Empty<int>());
                            lastRowOnPreviousPage = new Option<int>();
                            row = 0;
                        }
                        else
                        {
                            var endMergedRow = EndMergedRow(tableInfo.Table, mergedRows, row);
                            result.Add(Range(start, endMergedRow - start));
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
                    result.Add(Range(start, tableInfo.Table.Rows.Count - start));
            }
            endY = y;
            return result;
        }

        private static void Draw(TableInfo info, IEnumerable<int> rows, double y0, XGraphics xGraphics, Document document,
            Dictionary<Table, TableInfo> tableInfos, double x0, SyncPageInfo syncPageInfo, int tableLevel, Drawer drawer)
        {
            var firstRow = rows.FirstOrNone();
            if (!firstRow.HasValue) return;
            var maxTopBorder = firstRow.Value == 0
                ? MaxTopBorder(info)
                : MaxBottomBorder(firstRow.Value - 1, info.Table, info.BottomBorderFunc);
            var tableY = y0 + info.Table.TopMargin.ValueOr(0);
            {
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var topBorder = firstRow.Value == 0
                        ? info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index))
                        : info.BottomBorderFunc(new CellInfo(firstRow.Value - 1, column.Index));
                    if (topBorder.HasValue)
                    {
                        var borderY = tableY + maxTopBorder - topBorder.Value/2;
                        var leftBorder = column.Index == 0
                            ? info.LeftBorderFunc(new CellInfo(firstRow.Value, 0)).ValueOr(0)
                            : info.RightBorderFunc(new CellInfo(firstRow.Value, column.Index - 1)).ValueOr(0);
                        drawer.DrawLine(new XPen(XColors.Black, topBorder.Value),
                            x - leftBorder, borderY,
                            x + column.Width, borderY);
                    }
                    x += column.Width;
                }
            }
            var y = tableY + maxTopBorder;
            foreach (var row in rows)
            {
                var leftBorder = info.LeftBorderFunc(new CellInfo(row, 0));
                if (leftBorder.HasValue)
                {
                    var borderX = x0 + info.MaxLeftBorder - leftBorder.Value/2;
                    drawer.DrawLine(new XPen(XColors.Black, leftBorder.Value),
                        borderX, y, borderX, y + info.MaxHeights[row]);
                }
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var cell = info.Table.Rows[row].Cells[column.Index];
                    syncPageInfo.CellInfos.Add(new SyncCellInfo {
                        X = x,
                        Y = y,
                        Height = Range(0, cell.Rowspan().ValueOr(1)).Sum(i => info.MaxHeights[row + i]),
                        Width = Range(0, cell.Colspan().ValueOr(1)).Sum(i => info.Table.Columns[column.Index + i].Width),
                        CallerInfos = cell.CallerInfos,
                        TableLevel = tableLevel
                    });
                    var bottomBorder = info.BottomBorderFunc(new CellInfo(row, column.Index));
                    var backgroundColor = info.BackgroundColor(new CellInfo(cell));
                    if (backgroundColor.HasValue)
                        drawer.DrawRectangle(new XSolidBrush(backgroundColor.Value), x, y,
                            column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).ValueOr(0),
                            info.MaxHeights[row] - bottomBorder.ValueOr(0),
                            DrawType.Background);
                    HighlightCells(xGraphics, info, bottomBorder, row, column, x, y, tableY, drawer, document);
                    HighlightCellLine(xGraphics, info, bottomBorder, row, column, x, y, document, drawer);
                    var cellInnerHeight = Range(0, cell.Rowspan().ValueOr(1))
                        .Sum(i => info.MaxHeights[row + i] -
                            MaxBottomBorder(row + cell.Rowspan().ValueOr(1) - 1, info.Table, info.BottomBorderFunc));
                    var width = info.Table.ContentWidth(row, column, info.RightBorderFunc);
                    var contentHeight = cell.Elements.Sum(_ => _.Match(
                        p => p.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc),
                        t => t.GetTableHeight(xGraphics, tableInfos)));
                    double dy;
                    switch (cell.VerticalAlign())
                    {
                        case VerticalAlign.Top:
                            dy = 0;
                            break;
                        case VerticalAlign.Center:
                            dy = (cellInnerHeight - contentHeight) / 2;
                            break;
                        case VerticalAlign.Bottom:
                            dy = cellInnerHeight - contentHeight;
                            break;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                    double paragraphY = 0;
                    foreach (var element in cell.Elements)
                    {
                        if (document.ParagraphsAreHighlighted)
                            element.Match(
                                paragraph => {
                                    HighlightParagraph(paragraph, column, row, x, y + dy + paragraphY, width, info, xGraphics, drawer);
                                    return new { };
                                },
                                table => new { });
                        element.Match(
                            paragraph => {
                                ParagraphRenderer.Draw(xGraphics, paragraph, x, y + dy + paragraphY, width, paragraph.Alignment(), drawer);
                                return new { };
                            },
                            table => {
                                var tableInfo = GetTableInfo(tableInfos, xGraphics).GetValue(table);
                                double dx;
                                switch (table.HorizontalAlign())
                                {
                                    case HorizontalAlign.Left:
                                        dx = 0;
                                        break;
                                    case HorizontalAlign.Center:
                                        dx = (width - TableWidth(tableInfo)) / 2;
                                        break;
                                    case HorizontalAlign.Right:
                                        dx = width - TableWidth(tableInfo);
                                        break;
                                    default:
                                        throw new ArgumentOutOfRangeException();
                                }
                                Draw(tableInfo,
                                    Range(0, table.Rows.Count), y + dy + paragraphY, xGraphics,
                                    document, tableInfos, x + table.LeftMargin.ValueOr(0) + dx, syncPageInfo, tableLevel + 1, drawer);
                                return new { };
                            });
                        paragraphY += element.Match(
                            paragraph => paragraph.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc),
                            table => table.GetTableHeight(xGraphics, tableInfos));
                    }
                    var rightBorder = info.RightBorderFunc(new CellInfo(row, column.Index));
                    if (rightBorder.HasValue)
                    {
                        var borderX = x + column.Width - rightBorder.Value/2;
                        drawer.DrawLine(new XPen(XColors.Black, rightBorder.Value),
                            borderX, y, borderX, y + info.MaxHeights[row]);
                    }
                    if (bottomBorder.HasValue)
                    {
                        var borderY = y + info.MaxHeights[row] - bottomBorder.Value/2;
                        drawer.DrawLine(new XPen(XColors.Black, bottomBorder.Value),
                            x, borderY, x + column.Width, borderY);
                    }
                    x += column.Width;
                }
                y += info.MaxHeights[row];
            }
        }

        private static double MaxTopBorder(TableInfo info)
        {
            return info.Table.Columns.Max(column => info.TopBorderFunc(new CellInfo(0, column.Index)).ValueOr(0));
        }

        private static double TableWidth(TableInfo tableInfo)
        {
            return tableInfo.MaxLeftBorder + tableInfo.Table.Columns.Sum(_ => _.Width) +
                tableInfo.Table.LeftMargin.ValueOr(0) + tableInfo.Table.RightMargin.ValueOr(0);
        }

        private static double MaxBottomBorder(int rowIndex, Table table, Func<CellInfo, Option<double>> bottomBorderFunc)
        {
            return table.Columns.Max(column => bottomBorderFunc(new CellInfo(rowIndex, column.Index)).ValueOr(0));
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
                    var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan());
                    if (rowspan.HasValue)
                        for (var i = row.Index + 1; i < row.Index + rowspan.Value; i++)
                            set.Add(i);
                }
            return set;
        }

        private static double ContentWidth(this Table table, int row, Column column, Func<CellInfo, Option<double>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Colspan()).Match(
                colspan => column.Width
                    + Range(column.Index + 1, colspan - 1).Sum(i => table.Columns[i].Width)
                    - table.BorderWidth(row, column, column.Index + colspan - 1, rightBorderFunc),
                () => column.Width - table.BorderWidth(row, column, column.Index, rightBorderFunc));

        private static double BorderWidth(this Table table, int row, Column column, int rightColumn, Func<CellInfo, Option<double>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Rowspan()).Match(
                rowspan => Range(row, rowspan)
                    .Max(i => rightBorderFunc(new CellInfo(i, rightColumn)).ValueOr(0)),
                () => rightBorderFunc(new CellInfo(row, rightColumn)).ValueOr(0));

        private static Dictionary<int, double> MaxHeights(this Table table, XGraphics graphics, Func<CellInfo, Option<double>> rightBorderFunc,
            Func<CellInfo, Option<double>> bottomBorderFunc, Dictionary<Table, TableInfo> tableInfos)
        {
            var cellContentsByBottomRow = new Dictionary<CellInfo, MaxHeightsTuple>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var cell = table.Rows[row.Index].Cells[column.Index];
                    var elements = cell.Elements;
                    if (elements.Count > 0)
                    {
                        var rowspan = cell.Rowspan();
                        var rowIndex = rowspan.Match(value => row.Index + value - 1, () => row.Index);
                        cellContentsByBottomRow.Add(new CellInfo(rowIndex, column.Index),
                            new MaxHeightsTuple(elements, rowspan, row));
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
                        var paragraphHeight = cell.Elements.Sum(_ => _.Match(
                            p => p.GetParagraphHeight(cell.Row.Index, column, table, graphics, rightBorderFunc),
                            t => t.GetTableHeight(graphics, tableInfos)));
                        rowHeightByContent = cell.Rowspan.Match(
                            _ => Math.Max(paragraphHeight - Range(1, _ - 1).Sum(i => result[row.Index - i]), 0),
                            () => paragraphHeight);
                    }
                    else
                        rowHeightByContent = 0;
                    var innerHeight = row.Height().Match(
                        _ => Math.Max(rowHeightByContent, _), () => rowHeightByContent);
                    var height = innerHeight + MaxBottomBorder(row.Index, table, bottomBorderFunc);
                    if (maxHeight < height)
                        maxHeight = height;
                }
                result.Add(row.Index, maxHeight);
            }
            return result;
        }

        private static double GetTableHeight(this Table table, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos)
        {
            var tableInfo = GetTableInfo(tableInfos, graphics).GetValue(table);
            return MaxTopBorder(tableInfo) + table.Rows.Sum(row => tableInfo.MaxHeights[row.Index]) + table.TopMargin.ValueOr(0) +
                table.BottomMargin.ValueOr(0);
        }

        private static Lazy<Table, TableInfo> GetTableInfo(Dictionary<Table, TableInfo> tableInfos, XGraphics graphics)
        {
            return Lazy.Create(tableInfos, table => GetTableInfo(graphics, table, tableInfos));
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
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                        var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan());
                        if (rowspan.HasValue)
                            for (var i = 1; i <= rowspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                                    new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                    }
                    var leftBorder = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.LeftBorder);
                    if (leftBorder.HasValue)
                    {
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
                        var rowspan = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.Rowspan());
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
                    var bottomBorder = row.Cells[column.Index].BottomBorder;
                    if (bottomBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index), 
                            new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                        var colspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan());
                        if (colspan.HasValue)
                            for (var i = 1; i <= colspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                                    new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                    }
                    var topBorder = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.TopBorder);
                    if (topBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index),
                            new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
                        var colspan = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.Colspan());
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
                    var rowspan = table.Find(new CellInfo(row.Index, 0)).SelectMany(_ => _.Rowspan());
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
                    var colspan = table.Find(new CellInfo(0, column.Index)).SelectMany(_ => _.Colspan());
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

        private static Func<CellInfo, Option<XColor>> BackgroundColor(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BackgroundTuple>>();
            foreach (var row in table.Rows)
            foreach (var column in table.Columns)
            {
                var cell = row.Cells[column.Index];
                if (cell.BackgroundColor().HasValue)
                    for (var i = 0; i < cell.Rowspan().ValueOr(1); i++)
                    for (var j = 0; j < cell.Colspan().ValueOr(1); j++)
                        result.Add(new CellInfo(row.Index + i, column.Index + j),
                            new BackgroundTuple(cell.BackgroundColor().Value, new CellInfo(cell)));
            }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The background color is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToSttring(table)}");
                else
                    return list[0].Color;
            });
        }

        private class BackgroundTuple
        {
            public XColor Color { get; }
            public CellInfo CellInfo { get; }

            public BackgroundTuple(XColor color, CellInfo cellInfo)
            {
                Color = color;
                CellInfo = cellInfo;
            }
        }

        private static string CellsToSttring(this IEnumerable<CellInfo> cells, Table table) 
            => string.Join(", ", cells.Select(_ => $"L{table.Line} r{_.RowIndex + 1}c{_.ColumnIndex + 1}"));

        private static TableInfo GetTableInfo(XGraphics xGraphics, Table table, Dictionary<Table, TableInfo> tableInfos)
        {
            var rightBorderFunc = table.RightBorder();
            var bottomBorderFunc = table.BottomBorder();
            return new TableInfo(table, table.TopBorder(), bottomBorderFunc,
                table.MaxHeights(xGraphics, rightBorderFunc, bottomBorderFunc, tableInfos), table.LeftBorder(),
                rightBorderFunc, table.BackgroundColor());
        }

        private class TablePart
        {
            private double TableY { get; }
            public IEnumerable<int> Rows { get; }
            private int Index { get; }
            public TableInfo TableInfo { get; }
            public bool IsFirst => Index == 0;
            public double Y(PageSettings pageSettings) => IsFirst ? TableY : pageSettings.TopMargin;

            public TablePart(IEnumerable<int> rows, int index, TableInfo tableInfo, double tableY)
            {
                TableY = tableY;
                Rows = rows;
                Index = index;
                TableInfo = tableInfo;
            }
        }

        private class TableInfo
        {
            public Table Table { get; }
            public Func<CellInfo, Option<double>> RightBorderFunc { get; }
            public Func<CellInfo, Option<XColor>> BackgroundColor { get; }
            public Func<CellInfo, Option<double>> LeftBorderFunc { get; }
            public Func<CellInfo, Option<double>> TopBorderFunc { get; }
            public Func<CellInfo, Option<double>> BottomBorderFunc { get; }
            public Dictionary<int, double> MaxHeights { get; }
            public double MaxLeftBorder { get; }

            public TableInfo(Table table, Func<CellInfo, Option<double>> topBorderFunc, Func<CellInfo, Option<double>> bottomBorderFunc,
                Dictionary<int, double> maxHeights, Func<CellInfo, Option<double>> leftBorderFunc, Func<CellInfo, Option<double>> rightBorderFunc,
                Func<CellInfo, Option<XColor>> backgroundColor)
            {
                Table = table;
                RightBorderFunc = rightBorderFunc;
                BackgroundColor = backgroundColor;
                LeftBorderFunc = leftBorderFunc;
                TopBorderFunc = topBorderFunc;
                BottomBorderFunc = bottomBorderFunc;
                MaxHeights = maxHeights;
                MaxLeftBorder = table.Rows.Count == 0 ? 0 : table.Rows.Max(row => leftBorderFunc(new CellInfo(row.Index, 0)).ValueOr(0));
            }
        }

        private class MaxHeightsTuple
        {
            public List<IElement> Elements { get; }
            public Option<int> Rowspan { get; }
            public Row Row { get; }

            public MaxHeightsTuple(List<IElement> elements, Option<int> rowspan, Row row)
            {
                Elements = elements;
                Rowspan = rowspan;
                Row = row;
            }
        }

        private class BorderTuple
        {
            public double Width { get; }
            public CellInfo CellInfo { get; }

            public BorderTuple(double width, CellInfo cellInfo)
            {
                Width = width;
                CellInfo = cellInfo;
            }
        }

        private static void HighlightParagraph(Paragraph paragraph, Column column, int row, double x, double y, double width, TableInfo info, XGraphics xGraphics, Drawer drawer)
        {
            var innerHeight = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc);
            var innerWidth = paragraph.GetInnerWidth(width);
            if (innerWidth > 0 && innerHeight > 0)
                FillRectangle(drawer, XColor.FromArgb(32, 0, 0, 255),
                    x + paragraph.LeftMargin.ValueOr(0),
                    y + paragraph.TopMargin.ValueOr(0),
                    innerWidth, innerHeight);
        }

        private static void HighlightCells(XGraphics xGraphics, TableInfo info, Option<double> bottomBorder, int row, Column column, double x, double y,
            double tableY, Drawer drawer, Document document)
        {
            if (document.CellsAreHighlighted)
            {
                var color = (row + column.Index) % 2 == 1
                    ? XColor.FromArgb(32, 127, 127, 127)
                    : XColor.FromArgb(32, 0, 255, 0);
                var height = info.MaxHeights[row] - bottomBorder.ValueOr(0);
                var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).ValueOr(0);
                FillRectangle(drawer, color, x, y, width, height);
            }
            if (document.R1C1AreVisible)
            {
                var font = new XFont("Times New Roman", 10, XFontStyle.Regular,
                    new XPdfFontOptions(PdfFontEncoding.Unicode));
                var redBrush = new XSolidBrush(XColor.FromArgb(128, 255, 0, 0));
                var purpleBrush = new XSolidBrush(XColor.FromArgb(128, 87, 0, 127));
                if (column.Index == 0)
                {
                    var text = $"r{row + 1}";
                    var lineSpace = font.LineSpace(xGraphics);
                    var rnHeight = lineSpace * font.FontFamily.GetCellAscent(font.Style) / font.FontFamily.GetLineSpacing(font.Style);
                    var rnX = x - xGraphics.MeasureString(text, font).Width;
                    drawer.DrawString(text, font, redBrush, rnX, y + rnHeight);
                    if (row == 0)
                    {
                        var lineText = $"{info.Table.Line}";
                        drawer.DrawString(lineText,
                            font,
                            purpleBrush,
                            rnX - xGraphics.MeasureString($"{lineText} ", font, ParagraphRenderer.MeasureTrailingSpacesStringFormat).Width,
                            y + rnHeight);
                    }
                }
                if (row == 0)
                    drawer.DrawString($"c{column.Index + 1}", font, redBrush, x, tableY - 1);
            }
        }

        private static void HighlightCellLine(XGraphics xGraphics, TableInfo info, Option<double> bottomBorder, int row, Column column, double x, double y, Document document,
            Drawer drawer)
        {
            if (!document.CellLineNumbersAreVisible) return;
            var cell = info.Table.Rows[row].Cells[column.Index];
            if (cell.CallerInfos.Count <= 0) return;
            var text = string.Join(" ", cell.CallerInfos.Select(_ => _.Line));
            var font = new XFont("Arial", 7, XFontStyle.Regular,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            var height = info.MaxHeights[row] - bottomBorder.ValueOr(0);
            var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).ValueOr(0);
            drawer.DrawString(text,
                font,
                new XSolidBrush(XColor.FromArgb(128, 0, 0, 255)),
                x + width - xGraphics.MeasureString(text, font).Width,
                y + height);
        }

        private static void FillRectangle(Drawer drawer, XColor color, double x, double y, double width, double height)
        {
            var lineY = y + height / 2;
            //DrawRectangle does not draw the last pixel sometimes
            drawer.DrawLine(new XPen(color, height), x, lineY, x + width, lineY);
        }
    }
}