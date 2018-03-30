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
        public static List<SyncPageInfo> Draw(XGraphics xGraphics, Section section, Action<int, Action<XGraphics>> pageAction, IEnumerable<Table> tables,
            Document document, GraphicsType graphicsType)
        {
            var tableInfos = new Dictionary<Table, TableInfo>();
            var firstOnPage = true;
            var y = section.TopMargin(xGraphics, tableInfos, new TextMode.Measure(), document);
            var tableParts = tables.Select(table => {
                var tableInfo = GetTableInfo(tableInfos, xGraphics, new TextMode.Measure(), document).GetValue(table);
                var tableY = y;
                var rowSets = SplitByPages(tableInfo, firstOnPage, out var endY, section, tableY, xGraphics, tableInfos, new TextMode.Measure(), document);
                if (rowSets.Count > 0)
                    firstOnPage = false;
                y = endY + table.BottomMargin.ValueOr(0);
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
                void DrawHeaders(Drawer drawer, int pageIndex)
                {
                    foreach (var header in section.Headers)
                        Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document).GetValue(header),
                            Range(0, header.Rows.Count).Select(_ => new SplitedRow(_)), 0, xGraphics, document, tableInfos,
                            section.PageSettings.LeftMargin, syncPageInfo, 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count));                    
                }
                void DrawFooters(Drawer drawer, int pageIndex)
                {
                    foreach (var footer in section.Footers)
                        Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document).GetValue(footer),
                            Range(0, footer.Rows.Count).Select(_ => new SplitedRow(_)),
                            section.PageSettings.PageHeight - footer.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document),
                            xGraphics, document, tableInfos,
                            section.PageSettings.LeftMargin, syncPageInfo, 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count));                    
                }
                if (index == 0)
                {
                    var drawer = new Drawer(xGraphics);
                    DrawHeaders(drawer, index);
                    foreach (var part in pages[index])
                        Draw(part.TableInfo, part.Rows, part.Y(section, xGraphics, tableInfos, new TextMode.Draw(index, pages.Count), document), xGraphics, document, tableInfos,
                            section.PageSettings.LeftMargin, syncPageInfo, 0, drawer, graphicsType, new TextMode.Draw(index, pages.Count));
                    DrawFooters(drawer, index);
                    drawer.Flush();
                }
                else
                    pageAction(index, xGraphics2 => {
                        var drawer = new Drawer(xGraphics2);
                        DrawHeaders(drawer, index);
                        foreach (var part in pages[index])
                            Draw(part.TableInfo, part.Rows, part.Y(section, xGraphics, tableInfos, new TextMode.Draw(index, pages.Count), document), xGraphics2, document, tableInfos,
                                section.PageSettings.LeftMargin, syncPageInfo, 0, drawer, graphicsType, new TextMode.Draw(index, pages.Count));
                        DrawFooters(drawer, index);
                        drawer.Flush();
                    });
            }
            return syncPageInfos;
        }

        private static List<IEnumerable<SplitedRow>> SplitByPages(TableInfo tableInfo, bool firstOnPage, out double endY, Section section, double tableY,
            XGraphics xGraphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            if (tableInfo.Table.Rows.Count == 0)
            {
                endY = tableY;
                return new List<IEnumerable<SplitedRow>>();
            }
            var mergedRows = MergedRows(tableInfo.Table);
            var y = tableY + tableInfo.Table.TopMargin.ValueOr(0) +
                tableInfo.Table.Columns.Max(column => tableInfo.TopBorderFunc(new CellInfo(0, column.Index))
                    .Select(_ => _.Width).ValueOr(0));
            var lastRowOnPreviousPage = new Option<int>();
            var row = 0;
            var tableFirstPage = true;
            var result = new List<IEnumerable<SplitedRow>>();
            IEnumerable<SplitedRow> RowRange(int start, int count) =>
                (tableFirstPage ? Empty<SplitedRow>() : tableInfo.TableHeaderRows.OrderBy(_ => _).Select(_ => new SplitedRow(_)))
					.Concat(Range(start, count).Select(_ => new SplitedRow(_)));
            while (true)
            {
                y += tableInfo.MaxHeights[row];
                if (section.PageSettings.PageHeight - section.BottomMargin(xGraphics, tableInfos, mode, document) - y < 0)
                {
                    var firstMergedRow = FirstMergedRow(mergedRows, row);
                    var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
                    if (firstMergedRow - start > 0)
                    {
                        result.Add(RowRange(start, firstMergedRow - start));
                        lastRowOnPreviousPage = firstMergedRow - 1;
                        row = firstMergedRow;
                    }
                    else
                    {
                        if (firstMergedRow == 0 && tableFirstPage && !firstOnPage)
                        {
                            result.Add(Empty<SplitedRow>());
                            lastRowOnPreviousPage = new Option<int>();
                            row = 0;
                        }
                        else
                        {
                            var endMergedRow = EndMergedRow(tableInfo.Table, mergedRows, row);
                            result.Add(RowRange(start, endMergedRow - start));
                            lastRowOnPreviousPage = endMergedRow;
                            row = endMergedRow + 1;
                            if (row >= tableInfo.Table.Rows.Count) break;
                        }
                    }
                    tableFirstPage = false;
                    double topIndent;
                    if (tableInfo.TableHeaderRows.Count == 0)
                        topIndent = row == 0
                            ? tableInfo.Table.Columns.Max(column => tableInfo.TopBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0))
                            : tableInfo.Table.Columns.Max(column =>
                                tableInfo.BottomBorderFunc(new CellInfo(row - 1, column.Index)).Select(_ => _.Width).ValueOr(0));
                    else
                    {
                        var firstHeaderRow = tableInfo.TableHeaderRows.OrderBy(_ => _).First();
                        topIndent = (firstHeaderRow == 0
                                ? tableInfo.Table.Columns.Max(column =>
                                    tableInfo.TopBorderFunc(new CellInfo(firstHeaderRow, column.Index)).Select(_ => _.Width).ValueOr(0))
                                : tableInfo.Table.Columns.Max(column =>
                                    tableInfo.BottomBorderFunc(new CellInfo(firstHeaderRow - 1, column.Index)).Select(_ => _.Width).ValueOr(0))) +
                            tableInfo.TableHeaderRows.Sum(rowIndex => tableInfo.MaxHeights[rowIndex]);
                    }
                    y = section.TopMargin(xGraphics, tableInfos, mode, document) + topIndent;
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
                    result.Add(RowRange(start, tableInfo.Table.Rows.Count - start));
            }
            endY = y;
            return result;
        }

        private static void Draw(TableInfo info, IEnumerable<SplitedRow> rows, double y0, XGraphics xGraphics, Document document,
            Dictionary<Table, TableInfo> tableInfos, double x0, SyncPageInfo syncPageInfo, int tableLevel, Drawer drawer, GraphicsType graphicsType,
            TextMode mode)
        {
            var firstRow = rows.FirstOrNone();
            if (!firstRow.HasValue) return;
            var maxTopBorder = firstRow.Value.RowIndex == 0
                ? MaxTopBorder(info)
                : MaxBottomBorder(firstRow.Value.RowIndex - 1, info.Table, info.BottomBorderFunc);
            var tableY = y0 + info.Table.TopMargin.ValueOr(0);
            {
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var topBorder = firstRow.Value.RowIndex == 0
                        ? info.TopBorderFunc(new CellInfo(firstRow.Value.RowIndex, column.Index))
                        : info.BottomBorderFunc(new CellInfo(firstRow.Value.RowIndex - 1, column.Index));
                    if (topBorder.HasValue)
                    {
                        var borderY = tableY + maxTopBorder - topBorder.Value.Width/2;
                        var leftBorder = column.Index == 0
                            ? info.LeftBorderFunc(new CellInfo(firstRow.Value.RowIndex, 0)).Select(_ => _.Width).ValueOr(0)
                            : info.RightBorderFunc(new CellInfo(firstRow.Value.RowIndex, column.Index - 1)).Select(_ => _.Width).ValueOr(0);
                        drawer.DrawLine(topBorder.Value,
                            x - leftBorder, borderY,
                            x + column.Width, borderY);
                    }
                    x += column.Width;
                }
            }
            var y = tableY + maxTopBorder;
            foreach (var row in rows)
            {
                {
                    var leftBorder = info.LeftBorderFunc(new CellInfo(row.RowIndex, 0));
                    if (leftBorder.HasValue)
                    {
                        var borderX = x0 + info.MaxLeftBorder - leftBorder.Value.Width / 2;
                        drawer.DrawLine(leftBorder.Value,
                            borderX, y, borderX, y + info.MaxHeights[row.RowIndex]);
                    }
                }
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var cell = info.Table.Rows[row.RowIndex].Cells[column.Index];
                    syncPageInfo.ItemInfos.Add(new SyncItemInfo {
                        X = x,
                        Y = y,
                        Height = Range(0, cell.Rowspan().ValueOr(1)).Sum(i => info.MaxHeights[row.RowIndex + i]),
                        Width = Range(0, cell.Colspan().ValueOr(1)).Sum(i => info.Table.Columns[column.Index + i].Width),
                        CallerInfos = cell.CallerInfos,
                        TableLevel = tableLevel,
                        Level = 0
                    });
                    var bottomBorder = info.BottomBorderFunc(new CellInfo(row.RowIndex, column.Index));
                    var backgroundColor = info.BackgroundColor(new CellInfo(cell));
                    if (backgroundColor.HasValue)
                        drawer.DrawRectangle(new XSolidBrush(backgroundColor.Value), x, y,
                            column.Width - info.RightBorderFunc(new CellInfo(row.RowIndex, column.Index)).Select(_ => _.Width).ValueOr(0),
                            info.MaxHeights[row.RowIndex] - bottomBorder.Select(_ => _.Width).ValueOr(0),
                            DrawType.Background);
                    HighlightCells(xGraphics, info, bottomBorder, row.RowIndex, column, x, y, tableY, drawer, document);
                    HighlightCellLine(xGraphics, info, bottomBorder, row.RowIndex, column, x, y, document, drawer);
                    var cellInnerHeight = Range(0, cell.Rowspan().ValueOr(1))
                        .Sum(i => info.MaxHeights[row.RowIndex + i] -
                            MaxBottomBorder(row.RowIndex + cell.Rowspan().ValueOr(1) - 1, info.Table, info.BottomBorderFunc));
                    var width = info.Table.ContentWidth(row.RowIndex, column, info.RightBorderFunc);
                    var contentHeight = cell.Elements.Sum(_ => _.Match(
                        p => p.GetParagraphHeight(row.RowIndex, column, info.Table, xGraphics, info.RightBorderFunc, mode, document),
                        t => t.GetTableHeight(xGraphics, tableInfos, mode, document)));
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
                                    HighlightParagraph(paragraph, column, row.RowIndex, x, y + dy + paragraphY, width, info, xGraphics, drawer, mode, document);
                                    return new { };
                                },
                                table => new { });
                        element.Match(
                            paragraph => {
                                syncPageInfo.ItemInfos.Add(new SyncItemInfo {
                                    X = x,
                                    Y = y + dy + paragraphY,
                                    Height = paragraph.GetInnerHeight(xGraphics, info.Table, row.RowIndex, column, info.RightBorderFunc, mode, document),
                                    Width = paragraph.GetInnerWidth(width),
                                    CallerInfos = paragraph.CallerInfos,
                                    TableLevel = tableLevel,
                                    Level = 1
                                });
                                ParagraphRenderer.Draw(xGraphics, paragraph, x, y + dy + paragraphY, width, paragraph.Alignment(), drawer, graphicsType, mode,
									document);
                                return new { };
                            },
                            table => {
                                var tableInfo = GetTableInfo(tableInfos, xGraphics, mode, document).GetValue(table);
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
                                    Range(0, table.Rows.Count).Select(_ => new SplitedRow(_)), y + dy + paragraphY, xGraphics,
                                    document, tableInfos, x + table.LeftMargin.ValueOr(0) + dx, syncPageInfo, tableLevel + 1, drawer, graphicsType, mode);
                                return new { };
                            });
                        paragraphY += element.Match(
                            paragraph => paragraph.GetParagraphHeight(row.RowIndex, column, info.Table, xGraphics, info.RightBorderFunc, mode, document),
                            table => table.GetTableHeight(xGraphics, tableInfos, mode, document));
                    }
                    var rightBorder = info.RightBorderFunc(new CellInfo(row.RowIndex, column.Index));
                    if (rightBorder.HasValue)
                    {
                        double bottomBorderWidth;
                        if (!info.RightBorderFunc(new CellInfo(row.RowIndex + 1, column.Index)).HasValue &&
                            info.BottomBorderFunc(new CellInfo(row.RowIndex, column.Index + 1)).HasValue)
                            bottomBorderWidth = info.BottomBorderFunc(new CellInfo(row.RowIndex, column.Index)).Select(_ => _.Width).ValueOr(0);
                        else
                            bottomBorderWidth = 0d;
                        var borderX = x + column.Width - rightBorder.Value.Width/2;
                        drawer.DrawLine(rightBorder.Value,
                            borderX, y, borderX, y + info.MaxHeights[row.RowIndex] - bottomBorderWidth);
                    }
                    if (bottomBorder.HasValue)
                    {
                        var leftCell = new CellInfo(row.RowIndex, column.Index - 1);
                        var leftBorder = !info.RightBorderFunc(leftCell).HasValue &&
                            !info.BottomBorderFunc(leftCell).HasValue
                                ? info.RightBorderFunc(new CellInfo(row.RowIndex + 1, column.Index - 1))
                                    .Select(_ => _.Width).ValueOr(0)
                                : 0d;
                        double rightBorderWidth;
                        if (!info.BottomBorderFunc(new CellInfo(row.RowIndex, column.Index + 1)).HasValue)
                            rightBorderWidth = info.RightBorderFunc(new CellInfo(row.RowIndex, column.Index)).Select(_ => _.Width).ValueOr(0);
                        else
                            rightBorderWidth = 0d;
                        var borderY = y + info.MaxHeights[row.RowIndex] - bottomBorder.Value.Width/2;
                        drawer.DrawLine(bottomBorder.Value,
                            x - leftBorder, borderY, x + column.Width - rightBorderWidth, borderY);
                    }
                    x += column.Width;
                }
                y += info.MaxHeights[row.RowIndex];
            }
        }

        private static double TopMargin(this Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            return Math.Max(section.PageSettings.TopMargin,
                section.Headers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document)));
        }

        private static double BottomMargin(this Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            return Math.Max(section.PageSettings.BottomMargin,
                section.Footers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document)));
        }

        private static double MaxTopBorder(TableInfo info)
        {
            return info.Table.Columns.Max(column => info.TopBorderFunc(new CellInfo(0, column.Index)).Select(_ => _.Width).ValueOr(0));
        }

        private static double TableWidth(TableInfo tableInfo)
        {
            return tableInfo.MaxLeftBorder + tableInfo.Table.Columns.Sum(_ => _.Width) +
                tableInfo.Table.LeftMargin.ValueOr(0) + tableInfo.Table.RightMargin.ValueOr(0);
        }

        private static double MaxBottomBorder(int rowIndex, Table table, Func<CellInfo, Option<XPen>> bottomBorderFunc)
        {
            return table.Columns.Max(column => bottomBorderFunc(new CellInfo(rowIndex, column.Index)).Select(_ => _.Width).ValueOr(0));
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

        private static double ContentWidth(this Table table, int row, Column column, Func<CellInfo, Option<XPen>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Colspan()).Match(
                colspan => column.Width
                    + Range(column.Index + 1, colspan - 1).Sum(i => table.Columns[i].Width)
                    - table.BorderWidth(row, column, column.Index + colspan - 1, rightBorderFunc),
                () => column.Width - table.BorderWidth(row, column, column.Index, rightBorderFunc));

        private static double BorderWidth(this Table table, int row, Column column, int rightColumn, Func<CellInfo, Option<XPen>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Rowspan()).Match(
                rowspan => Range(row, rowspan)
                    .Max(i => rightBorderFunc(new CellInfo(i, rightColumn)).Select(_ => _.Width).ValueOr(0)),
                () => rightBorderFunc(new CellInfo(row, rightColumn)).Select(_ => _.Width).ValueOr(0));

        private static Dictionary<int, double> MaxHeights(this Table table, XGraphics graphics, Func<CellInfo, Option<XPen>> rightBorderFunc,
            Func<CellInfo, Option<XPen>> bottomBorderFunc, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
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
                    if (cellContentsByBottomRow.TryGetValue(new CellInfo(row, column), out var cell))
                    {
                        var paragraphHeight = cell.Elements.Sum(_ => _.Match(
                            p => p.GetParagraphHeight(cell.Row.Index, column, table, graphics, rightBorderFunc, mode, document),
                            t => t.GetTableHeight(graphics, tableInfos, mode, document)));
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

        private static double GetTableHeight(this Table table, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            var tableInfo = GetTableInfo(tableInfos, graphics, mode, document).GetValue(table);
            return MaxTopBorder(tableInfo) + table.Rows.Sum(row => tableInfo.MaxHeights[row.Index]) + table.TopMargin.ValueOr(0) +
                table.BottomMargin.ValueOr(0);
        }

        private static Lazy<Table, TableInfo> GetTableInfo(Dictionary<Table, TableInfo> tableInfos, XGraphics graphics, TextMode mode, Document document)
        {
            return Lazy.Create(tableInfos, table => GetTableInfo(graphics, table, tableInfos, mode, document));
        }

        private static double GetParagraphHeight(this Paragraph paragraph, int row, Column column, Table table, XGraphics graphics,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document)
        {
            return paragraph.GetInnerHeight(graphics, table, row, column, rightBorderFunc, mode, document) +
                paragraph.TopMargin.ValueOr(0) + paragraph.BottomMargin.ValueOr(0);
        }

        private static double GetInnerHeight(this Paragraph paragraph, XGraphics graphics, Table table, int row, Column column,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document)
        {
            return ParagraphRenderer.GetHeight(graphics, paragraph, table.ContentWidth(row, column, rightBorderFunc), mode, document);
        }

        private static Func<CellInfo, Option<XPen>> RightBorder(this Table table)
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
                    return list[0].Value;
            });
        }

        private static Func<CellInfo, Option<XPen>> BottomBorder(this Table table)
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
                    return list[0].Value;
            });
        }

        private static Func<CellInfo, Option<XPen>> LeftBorder(this Table table)
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
                    return list[0].Value;
            });
        }

        private static Func<CellInfo, Option<XPen>> TopBorder(this Table table)
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
                    return list[0].Value;
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

        private static TableInfo GetTableInfo(XGraphics xGraphics, Table table, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            var rightBorderFunc = table.RightBorder();
            var bottomBorderFunc = table.BottomBorder();
            return new TableInfo(table, table.TopBorder(), bottomBorderFunc,
                table.MaxHeights(xGraphics, rightBorderFunc, bottomBorderFunc, tableInfos, mode, document), table.LeftBorder(),
                rightBorderFunc, table.BackgroundColor(), TableHeaderRows(table));
        }

        private static HashSet<int> TableHeaderRows(Table table)
        {
            var hashSet = new HashSet<int>();
            foreach (var row in table.Rows)
            foreach (var column in table.Columns)
                for (var i = 0; i < row.Cells[column.Index].Rowspan().ValueOr(1); i++)
                    if (row.TableHeader())
                        hashSet.Add(row.Index);
            return hashSet;
        }

        private class TablePart
        {
            private double TableY { get; }
            public IEnumerable<SplitedRow> Rows { get; }
            private int Index { get; }
            public TableInfo TableInfo { get; }
            public bool IsFirst => Index == 0;
            public double Y(Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document) => 
                IsFirst ? TableY : section.TopMargin(graphics, tableInfos, mode, document);

            public TablePart(IEnumerable<SplitedRow> rows, int index, TableInfo tableInfo, double tableY)
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
            public Func<CellInfo, Option<XPen>> RightBorderFunc { get; }
            public Func<CellInfo, Option<XColor>> BackgroundColor { get; }
            public HashSet<int> TableHeaderRows { get; }
            public Func<CellInfo, Option<XPen>> LeftBorderFunc { get; }
            public Func<CellInfo, Option<XPen>> TopBorderFunc { get; }
            public Func<CellInfo, Option<XPen>> BottomBorderFunc { get; }
            public Dictionary<int, double> MaxHeights { get; }
            public double MaxLeftBorder { get; }

            public TableInfo(Table table, Func<CellInfo, Option<XPen>> topBorderFunc, Func<CellInfo, Option<XPen>> bottomBorderFunc,
                Dictionary<int, double> maxHeights, Func<CellInfo, Option<XPen>> leftBorderFunc, Func<CellInfo, Option<XPen>> rightBorderFunc,
                Func<CellInfo, Option<XColor>> backgroundColor, HashSet<int> tableHeaderRows)
            {
                Table = table;
                RightBorderFunc = rightBorderFunc;
                BackgroundColor = backgroundColor;
                TableHeaderRows = tableHeaderRows;
                LeftBorderFunc = leftBorderFunc;
                TopBorderFunc = topBorderFunc;
                BottomBorderFunc = bottomBorderFunc;
                MaxHeights = maxHeights;
                MaxLeftBorder = table.Rows.Count == 0
                    ? 0
                    : table.Rows.Max(row => leftBorderFunc(new CellInfo(row.Index, 0))
                        .Select(_ => _.Width).ValueOr(0));
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
            public XPen Value { get; }
            public CellInfo CellInfo { get; }

            public BorderTuple(XPen value, CellInfo cellInfo)
            {
                Value = value;
                CellInfo = cellInfo;
            }
        }

        private static void HighlightParagraph(Paragraph paragraph, Column column, int row, double x, double y, double width, TableInfo info, XGraphics xGraphics, Drawer drawer,
            TextMode mode, Document document)
        {
            var innerHeight = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc, mode, document);
            var innerWidth = paragraph.GetInnerWidth(width);
            if (innerWidth > 0 && innerHeight > 0)
                FillRectangle(drawer, XColor.FromArgb(32, 0, 0, 255),
                    x + paragraph.LeftMargin.ValueOr(0),
                    y + paragraph.TopMargin.ValueOr(0),
                    innerWidth, innerHeight);
        }

        private static void HighlightCells(XGraphics xGraphics, TableInfo info, Option<XPen> bottomBorder, int row, Column column, double x, double y,
            double tableY, Drawer drawer, Document document)
        {
            if (document.CellsAreHighlighted)
            {
                var color = (row + column.Index) % 2 == 1
                    ? XColor.FromArgb(32, 127, 127, 127)
                    : XColor.FromArgb(32, 0, 255, 0);
                var height = info.MaxHeights[row] - bottomBorder.Select(_ => _.Width).ValueOr(0);
                var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
                FillRectangle(drawer, color, x, y, width, height);
            }
            if (document.R1C1AreVisible)
            {
                var font = new XFont("Times New Roman", 10, XFontStyle.Regular,
                    new XPdfFontOptions(PdfFontEncoding.Unicode));
                var redBrush = new XSolidBrush(XColor.FromArgb(200, 255, 0, 0));
                var purpleBrush = new XSolidBrush(XColor.FromArgb(200, 87, 0, 127));
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

        private static void HighlightCellLine(XGraphics xGraphics, TableInfo info, Option<XPen> bottomBorder, int row, Column column, double x, double y, Document document,
            Drawer drawer)
        {
            if (!document.CellLineNumbersAreVisible) return;
            var cell = info.Table.Rows[row].Cells[column.Index];
            if (cell.CallerInfos.Count <= 0) return;
            var text = string.Join(" ", cell.CallerInfos.Select(_ => _.Line));
            var font = new XFont("Arial", 7, XFontStyle.Regular,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            var height = info.MaxHeights[row] - bottomBorder.Select(_ => _.Width).ValueOr(0);
            var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
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