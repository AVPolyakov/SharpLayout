using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static System.Linq.Enumerable;
using static System.Math;

namespace SharpLayout
{
    public static class TableRenderer
    {
	    private class TableGroup
	    {
		    public List<Table> Tables { get; }

		    public TableGroup(List<Table> tables)
		    {
			    Tables = tables;
		    }
	    }

	    private static IEnumerable<TableGroup> GetTableGroups(this IEnumerable<Table> tables)
	    {
		    var list = new List<Table>();
		    foreach (var table in tables)
		    {
			    list.Add(table);
			    if (!table.KeepWithNext.ToOption().ValueOr(false))
			    {
				    yield return new TableGroup(list);
				    list = new List<Table>();
			    }
		    }
		    yield return new TableGroup(list);
	    }

        public static List<SyncPageInfo> Draw(XGraphics xGraphics, Section section, Action<int, Action<XGraphics>> pageAction, IEnumerable<Table> tables,
            Document document, GraphicsType graphicsType)
        {
            var tableInfos = new Dictionary<Table, TableInfo>();
            var firstOnPage = true;
            var y = section.TopMargin(xGraphics, tableInfos, new TextMode.Measure(), document);
            var tableParts = tables.GetTableGroups().Select(tableGroup => {
                var tableY = y;
                var slices = SplitByPages(tableGroup, firstOnPage, out var endY, section, tableY, xGraphics, tableInfos, new TextMode.Measure(), document);
	            var parts = slices.SelectMany(slice => slice.Rows.Select((rows, index) => new TablePart(rows, index, slice.TableInfo, slice.TableY))).ToList();
	            if (parts.Count > 0)
                    firstOnPage = false;
                y = endY;
                return parts;
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
                            Range(0, header.Rows.Count), 0, xGraphics, document, tableInfos,
                            section.PageSettings.LeftMargin, syncPageInfo, 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count));                    
                }
                void DrawFooters(Drawer drawer, int pageIndex)
                {
                    foreach (var footer in section.Footers)
                        Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document).GetValue(footer),
                            Range(0, footer.Rows.Count),
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

	    private class TableSlice
	    {
		    public List<IEnumerable<int>> Rows { get; }
		    public double TableY { get; }
		    public TableInfo TableInfo { get; }

		    public TableSlice(List<IEnumerable<int>> rows, double tableY, TableInfo tableInfo)
		    {
			    Rows = rows;
			    TableY = tableY;
			    TableInfo = tableInfo;
		    }
	    }

        private static List<TableSlice> SplitByPages(TableGroup tableGroup, bool firstOnPage, out double endY, Section section, double tableY,
            XGraphics xGraphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
	        var slices = new List<TableSlice>();
	        var currentTableY = tableY;
	        var infos = tableGroup.Tables.Select(table => GetTableInfo(tableInfos, xGraphics, new TextMode.Measure(), document).GetValue(table)).ToList();
	        var tableGroupFirstPage = true;
	        for (var tableIndex = 0; tableIndex < infos.Count; tableIndex++)
	        {
		        double currentEntY;
		        if (infos[tableIndex].Table.Rows.Count == 0)
		        {
			        currentEntY = currentTableY + infos[tableIndex].Table.BottomMargin().ToOption().ValueOr(0);
			        slices.Add(new TableSlice(new List<IEnumerable<int>>(), currentTableY, infos[tableIndex]));
		        }
		        else
		        {
			        var mergedRows = MergedRows(infos[tableIndex].Table);
			        var keepWithRows = KeepWithRows(infos[tableIndex].Table);
			        var y = currentTableY + infos[tableIndex].Table.TopMargin().ToOption().ValueOr(0) +
				        infos[tableIndex].Table.Columns.Max(column => infos[tableIndex].TopBorderFunc(new CellInfo(0, column.Index))
					        .Select(_ => _.Width).ValueOr(0));
			        var lastRowOnPreviousPage = new Option<int>();
			        var row = 0;
			        var sliceRows = new List<IEnumerable<int>>();
			        var addHeader = false;
			        IEnumerable<int> RowRange(int start, int count) =>
				        (!addHeader ? Empty<int>() : infos[tableIndex].TableHeaderRows.OrderBy(_ => _)).Concat(Range(start, count));
			        while (true)
			        {
				        y += infos[tableIndex].MaxHeights[row];
				        if (section.PageSettings.PageHeight - section.BottomMargin(xGraphics, tableInfos, mode, document) - y < 0)
				        {
					        var firstMergedRow = Min(FirstdRow(mergedRows, row), FirstdRow(keepWithRows, row));
					        var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
					        if (firstMergedRow - start > 0)
					        {
						        sliceRows.Add(RowRange(start, firstMergedRow - start));
						        addHeader = true;
						        lastRowOnPreviousPage = firstMergedRow - 1;
						        row = firstMergedRow;
					        }
					        else
					        {
						        if (firstMergedRow == 0 && tableGroupFirstPage && !firstOnPage)
						        {
							        sliceRows.Add(Empty<int>());
							        lastRowOnPreviousPage = new Option<int>();
							        row = 0;
							        tableIndex = 0;
							        slices.Clear();
						        }
						        else
						        {
							        var endMergedRow = Max(EndRow(infos[tableIndex].Table, mergedRows, row), EndRow(infos[tableIndex].Table, keepWithRows, row));
							        sliceRows.Add(RowRange(start, endMergedRow - start + 1));
							        addHeader = true;
							        lastRowOnPreviousPage = endMergedRow;
							        row = endMergedRow + 1;
							        if (row >= infos[tableIndex].Table.Rows.Count) break;
						        }
					        }
					        tableGroupFirstPage = false;
					        double topIndent;
					        if (infos[tableIndex].TableHeaderRows.Count == 0)
						        topIndent = row == 0
							        ? infos[tableIndex].Table.Columns.Max(column => infos[tableIndex].TopBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0))
							        : infos[tableIndex].Table.Columns.Max(column =>
								        infos[tableIndex].BottomBorderFunc(new CellInfo(row - 1, column.Index)).Select(_ => _.Width).ValueOr(0));
					        else
					        {
						        var firstHeaderRow = infos[tableIndex].TableHeaderRows.OrderBy(_ => _).First();
						        topIndent = (firstHeaderRow == 0
								        ? infos[tableIndex].Table.Columns.Max(column =>
									        infos[tableIndex].TopBorderFunc(new CellInfo(firstHeaderRow, column.Index)).Select(_ => _.Width).ValueOr(0))
								        : infos[tableIndex].Table.Columns.Max(column =>
									        infos[tableIndex].BottomBorderFunc(new CellInfo(firstHeaderRow - 1, column.Index)).Select(_ => _.Width).ValueOr(0))) +
							        infos[tableIndex].TableHeaderRows.Sum(rowIndex => infos[tableIndex].MaxHeights[rowIndex]);
					        }
					        y = section.TopMargin(xGraphics, tableInfos, mode, document) + topIndent;
				        }
				        else
				        {
					        row++;
					        if (row >= infos[tableIndex].Table.Rows.Count) break;
				        }
			        }
			        {
				        var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
				        if (start < infos[tableIndex].Table.Rows.Count)
					        sliceRows.Add(RowRange(start, infos[tableIndex].Table.Rows.Count - start));
			        }
			        currentEntY = y + infos[tableIndex].Table.BottomMargin().ToOption().ValueOr(0);
			        slices.Add(new TableSlice(sliceRows, currentTableY, infos[tableIndex]));
		        }
		        currentTableY = currentEntY;
	        }
	        endY = currentTableY;
	        return slices;
        }

        private static void Draw(TableInfo info, IEnumerable<int> rows, double y0, XGraphics xGraphics, Document document,
            Dictionary<Table, TableInfo> tableInfos, double x0, SyncPageInfo syncPageInfo, int tableLevel, Drawer drawer, GraphicsType graphicsType,
            TextMode mode)
        {
            var firstRow = rows.FirstOrNone();
            if (!firstRow.HasValue) return;
            var maxTopBorder = firstRow.Value == 0
                ? MaxTopBorder(info)
                : MaxBottomBorder(firstRow.Value - 1, info.Table, info.BottomBorderFunc);
            var tableY = y0 + info.Table.TopMargin().ToOption().ValueOr(0);
            {
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var topBorder = firstRow.Value == 0
                        ? info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index))
                        : info.BottomBorderFunc(new CellInfo(firstRow.Value - 1, column.Index));
                    if (topBorder.HasValue)
                    {
                        var borderY = tableY + maxTopBorder - topBorder.Value.Width/2;
                        var leftBorder = column.Index == 0
                            ? info.LeftBorderFunc(new CellInfo(firstRow.Value, 0)).Select(_ => _.Width).ValueOr(0)
                            : info.RightBorderFunc(new CellInfo(firstRow.Value, column.Index - 1)).Select(_ => _.Width).ValueOr(0);
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
                    var leftBorder = info.LeftBorderFunc(new CellInfo(row, 0));
                    if (leftBorder.HasValue)
                    {
                        var borderX = x0 + info.MaxLeftBorder - leftBorder.Value.Width / 2;
                        drawer.DrawLine(leftBorder.Value,
                            borderX, y, borderX, y + info.MaxHeights[row]);
                    }
                }
                var x = x0 + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var cell = info.Table.Rows[row].Cells[column.Index];
                    if (graphicsType == GraphicsType.Image)
                        syncPageInfo.ItemInfos.Add(new SyncItemInfo {
                            X = x,
                            Y = y,
                            Height = Range(0, cell.Rowspan().ToOption().ValueOr(1)).Sum(i => info.MaxHeights[row + i]),
                            Width = Range(0, cell.Colspan().ToOption().ValueOr(1)).Sum(i => info.Table.Columns[column.Index + i].Width),
                            CallerInfos = cell.CallerInfos ?? new List<CallerInfo>(),
                            TableLevel = tableLevel,
                            Level = 0
                        });
                    var bottomBorder = info.BottomBorderFunc(new CellInfo(row, column.Index));
                    var backgroundColor = info.BackgroundColor(new CellInfo(cell));
                    if (backgroundColor.HasValue)
                        drawer.DrawRectangle(new XSolidBrush(backgroundColor.Value), x, y,
                            column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0),
                            info.MaxHeights[row] - bottomBorder.Select(_ => _.Width).ValueOr(0),
                            DrawType.Background);
                    HighlightCells(xGraphics, info, bottomBorder, row, column, x, y, tableY, drawer, document);
                    HighlightCellLine(xGraphics, info, bottomBorder, row, column, x, y, document, drawer);
                    var cellInnerHeight = Range(0, cell.Rowspan().ToOption().ValueOr(1))
                        .Sum(i => info.MaxHeights[row + i] -
                            MaxBottomBorder(row + cell.Rowspan().ToOption().ValueOr(1) - 1, info.Table, info.BottomBorderFunc));
                    var width = info.Table.ContentWidth(row, column, info.RightBorderFunc);
                    var contentHeight = cell.Elements.Sum(_ => _.Match(
                        p => p.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc, mode, document),
                        t => t.GetTableHeight(xGraphics, tableInfos, mode, document)));
                    double dy;
                    switch (GetVerticalAlign(cell, info.Table))
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
                                    HighlightParagraph(paragraph, column, row, x, y + dy + paragraphY, width, info, xGraphics, drawer, mode, document);
                                    return new { };
                                },
                                table => new { });
                        element.Match(
                            paragraph => {
                                if (graphicsType == GraphicsType.Image)
                                    syncPageInfo.ItemInfos.Add(new SyncItemInfo {
                                        X = x,
                                        Y = y + dy + paragraphY,
                                        Height = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc, mode, document),
                                        Width = paragraph.GetInnerWidth(width),
                                        CallerInfos = paragraph.CallerInfos ?? new List<CallerInfo>(),
                                        TableLevel = tableLevel,
                                        Level = 1
                                    });
                                ParagraphRenderer.Draw(xGraphics, paragraph, x, y + dy + paragraphY, width, Alignment(paragraph, info.Table), drawer, graphicsType, mode,
									document, info.Table);
                                return new { };
                            },
                            table => {
                                var tableInfo = GetTableInfo(tableInfos, xGraphics, mode, document).GetValue(table);
                                double dx;
                                switch (table.Alignment().ToOption().ValueOr(HorizontalAlign.Left))
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
                                    document, tableInfos, x + table.LeftMargin().ToOption().ValueOr(0) + dx, syncPageInfo, tableLevel + 1, drawer, graphicsType, mode);
                                return new { };
                            });
                        paragraphY += element.Match(
                            paragraph => paragraph.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc, mode, document),
                            table => table.GetTableHeight(xGraphics, tableInfos, mode, document));
                    }
                    var rightBorder = info.RightBorderFunc(new CellInfo(row, column.Index));
                    if (rightBorder.HasValue)
                    {
                        double bottomBorderWidth;
                        if (!info.RightBorderFunc(new CellInfo(row + 1, column.Index)).HasValue &&
                            info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue)
                            bottomBorderWidth = info.BottomBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
                        else
                            bottomBorderWidth = 0d;
                        var borderX = x + column.Width - rightBorder.Value.Width/2;
                        drawer.DrawLine(rightBorder.Value,
                            borderX, y, borderX, y + info.MaxHeights[row] - bottomBorderWidth);
                    }
                    if (bottomBorder.HasValue)
                    {
                        var leftCell = new CellInfo(row, column.Index - 1);
                        var leftBorder = !info.RightBorderFunc(leftCell).HasValue &&
                            !info.BottomBorderFunc(leftCell).HasValue
                                ? info.RightBorderFunc(new CellInfo(row + 1, column.Index - 1))
                                    .Select(_ => _.Width).ValueOr(0)
                                : 0d;
                        double rightBorderWidth;
                        if (!info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue)
                            rightBorderWidth = info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
                        else
                            rightBorderWidth = 0d;
                        var borderY = y + info.MaxHeights[row] - bottomBorder.Value.Width/2;
                        drawer.DrawLine(bottomBorder.Value,
                            x - leftBorder, borderY, x + column.Width - rightBorderWidth, borderY);
                    }
                    x += column.Width;
                }
                y += info.MaxHeights[row];
            }
        }

	    private static VerticalAlign GetVerticalAlign(Cell cell, Table table)
	    {
		    if (cell.VerticalAlign().HasValue) return cell.VerticalAlign().Value;
		    if (table.ContentVerticalAlign().HasValue) return table.ContentVerticalAlign().Value;
		    return VerticalAlign.Top;
	    }

	    private static HorizontalAlign Alignment(Paragraph paragraph, Table table)
	    {
		    if (paragraph.Alignment().HasValue) return paragraph.Alignment().Value;
		    if (table.ContentAlign().HasValue) return table.ContentAlign().Value;
			return HorizontalAlign.Left;
	    }

	    private static double TopMargin(this Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            return Max(section.PageSettings.TopMargin,
                section.Headers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document)));
        }

        private static double BottomMargin(this Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document)
        {
            return Max(section.PageSettings.BottomMargin,
                section.Footers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document)));
        }

        private static double MaxTopBorder(TableInfo info)
        {
            return info.Table.Columns.Max(column => info.TopBorderFunc(new CellInfo(0, column.Index)).Select(_ => _.Width).ValueOr(0));
        }

        private static double TableWidth(TableInfo tableInfo)
        {
            return tableInfo.MaxLeftBorder + tableInfo.Table.Columns.Sum(_ => _.Width) +
                tableInfo.Table.LeftMargin().ToOption().ValueOr(0) + tableInfo.Table.RightMargin().ToOption().ValueOr(0);
        }

        private static double MaxBottomBorder(int rowIndex, Table table, Func<CellInfo, Option<XPen>> bottomBorderFunc)
        {
            return table.Columns.Max(column => bottomBorderFunc(new CellInfo(rowIndex, column.Index)).Select(_ => _.Width).ValueOr(0));
        }

        private static int EndRow(Table table, HashSet<int> set, int row)
        {
            if (row + 1 >= table.Rows.Count) return row;
            var i = row + 1;
            while (true)
            {
                if (!set.Contains(i))
                    return i - 1;
                i++;
            }
        }

        private static int FirstdRow(HashSet<int> set, int row)
        {
            var i = row;
            while (true)
            {
                if (!set.Contains(i))
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
                    var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan().ToOption());
                    if (rowspan.HasValue)
                        for (var i = row.Index + 1; i < row.Index + rowspan.Value; i++)
                            set.Add(i);
                }
            return set;
        }

	    private static HashSet<int> KeepWithRows(Table table)
	    {
		    var set = new HashSet<int>();
		    foreach (var row in table.Rows)
		    {
			    var keepWith = row.KeepWith();
			    if (keepWith.HasValue)
				    for (var i = row.Index + 1; i < row.Index + keepWith.Value; i++)
					    set.Add(i);
		    }
		    return set;
	    }

        private static double ContentWidth(this Table table, int row, Column column, Func<CellInfo, Option<XPen>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Colspan().ToOption()).Match(
                colspan => column.Width
                    + Range(column.Index + 1, colspan - 1).Sum(i => table.Columns[i].Width)
                    - table.BorderWidth(row, column, column.Index + colspan - 1, rightBorderFunc),
                () => column.Width - table.BorderWidth(row, column, column.Index, rightBorderFunc));

        private static double BorderWidth(this Table table, int row, Column column, int rightColumn, Func<CellInfo, Option<XPen>> rightBorderFunc)
            => table.Find(new CellInfo(row, column.Index)).SelectMany(_ => _.Rowspan().ToOption()).Match(
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
                        var rowspan = cell.Rowspan().ToOption();
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
                            _ => Max(paragraphHeight - Range(1, _ - 1).Sum(i => result[row.Index - i]), 0),
                            () => paragraphHeight);
                    }
                    else
                        rowHeightByContent = 0;
                    var innerHeight = row.Height().ToOption().Match(
                        _ => Max(rowHeightByContent, _), () => rowHeightByContent);
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
            return MaxTopBorder(tableInfo) + table.Rows.Sum(row => tableInfo.MaxHeights[row.Index]) + table.TopMargin().ToOption().ValueOr(0) +
                table.BottomMargin().ToOption().ValueOr(0);
        }

        private static Lazy<Table, TableInfo> GetTableInfo(Dictionary<Table, TableInfo> tableInfos, XGraphics graphics, TextMode mode, Document document)
        {
            return Lazy.Create(tableInfos, table => GetTableInfo(graphics, table, tableInfos, mode, document));
        }

        private static double GetParagraphHeight(this Paragraph paragraph, int row, Column column, Table table, XGraphics graphics,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document)
        {
            return paragraph.GetInnerHeight(graphics, table, row, column, rightBorderFunc, mode, document) +
                paragraph.TopMargin().ToOption().ValueOr(0) + paragraph.BottomMargin().ToOption().ValueOr(0);
        }

        private static double GetInnerHeight(this Paragraph paragraph, XGraphics graphics, Table table, int row, Column column,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document)
        {
            return ParagraphRenderer.GetHeight(graphics, paragraph, table.ContentWidth(row, column, rightBorderFunc), mode, document, table);
        }

        private static Func<CellInfo, Option<XPen>> RightBorder(this Table table)
        {
            var rightMergedCells = RightMergedCells(table);
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var rightBorder = table.Find(new CellInfo(row, column)).SelectMany(_ => _.RightBorder());
                    if (rightBorder.HasValue)
                    {
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan().ToOption()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                        var rowspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan().ToOption());
                        if (rowspan.HasValue)
                            for (var i = 1; i <= rowspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                                    new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                    }
                    var leftBorder = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.LeftBorder());
                    if (leftBorder.HasValue)
                    {
                        var mergeRight = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan().ToOption()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                            new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
                        var rowspan = table.Find(new CellInfo(row.Index, column.Index + 1)).SelectMany(_ => _.Rowspan().ToOption());
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
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.Rows.Count) return new Option<XPen>();
                if (rightMergedCells.Contains(new CellInfo(cell.RowIndex, cell.ColumnIndex + 1))) return new Option<XPen>();
                return table.Border();
            });
        }

        private static HashSet<CellInfo> BottomMergedCells(Table table)
        {
            var set = new HashSet<CellInfo>();
            foreach (var row in table.Rows)
            foreach (var column in table.Columns)
            {
                var cell = table.Rows[row.Index].Cells[column.Index];
                var rowspan = cell.Rowspan().ToOption().ValueOr(1);
                var colspan = cell.Colspan().ToOption().ValueOr(1);
                for (var i = row.Index + 1; i < row.Index + rowspan; i++)
                for (var j = column.Index; j < column.Index + colspan; j++)
                    set.Add(new CellInfo(i, j));
            }
            return set;
        }

        private static HashSet<CellInfo> RightMergedCells(Table table)
        {
            var set = new HashSet<CellInfo>();
            foreach (var row in table.Rows)
            foreach (var column in table.Columns)
            {
                var cell = table.Rows[row.Index].Cells[column.Index];
                var rowspan = cell.Rowspan().ToOption().ValueOr(1);
                var colspan = cell.Colspan().ToOption().ValueOr(1);
                for (var i = row.Index; i < row.Index + rowspan; i++)
                for (var j = column.Index + 1; j < column.Index + colspan; j++)
                    set.Add(new CellInfo(i, j));
            }
            return set;
        }

        private static Func<CellInfo, Option<XPen>> BottomBorder(this Table table)
        {
            var bottomMergedCells = BottomMergedCells(table);
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
                foreach (var column in table.Columns)
                {
                    var bottomBorder = row.Cells[column.Index].BottomBorder();
                    if (bottomBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan().ToOption()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index), 
                            new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                        var colspan = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Colspan().ToOption());
                        if (colspan.HasValue)
                            for (var i = 1; i <= colspan.Value - 1; i++)
                                result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                                    new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                    }
                    var topBorder = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.TopBorder());
                    if (topBorder.HasValue)
                    {
                        var mergeDown = table.Find(new CellInfo(row, column)).SelectMany(_ => _.Rowspan().ToOption()).Match(_ => _ - 1, () => 0);
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index),
                            new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
                        var colspan = table.Find(new CellInfo(row.Index + 1, column.Index)).SelectMany(_ => _.Colspan().ToOption());
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
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.Rows.Count) return new Option<XPen>();
                if (bottomMergedCells.Contains(new CellInfo(cell.RowIndex + 1, cell.ColumnIndex))) return new Option<XPen>();
                return table.Border();
            });
        }

        private static Func<CellInfo, Option<XPen>> LeftBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var row in table.Rows)
            {
                var leftBorder = table.Find(new CellInfo(row.Index, 0)).SelectMany(_ => _.LeftBorder());
                if (leftBorder.HasValue)
                {
                    result.Add(new CellInfo(row.Index, 0), new BorderTuple(leftBorder.Value, new CellInfo(row.Index, 0)));
                    var rowspan = table.Find(new CellInfo(row.Index, 0)).SelectMany(_ => _.Rowspan().ToOption());
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
            }).Match(_ => _, table.Border);
        }

        private static Func<CellInfo, Option<XPen>> TopBorder(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var column in table.Columns)
            {
                var bottomBorder = table.Find(new CellInfo(0, column.Index)).SelectMany(_ => _.TopBorder());
                if (bottomBorder.HasValue)
                {
                    result.Add(new CellInfo(0, column.Index),
                        new BorderTuple(bottomBorder.Value, new CellInfo(0, column.Index)));
                    var colspan = table.Find(new CellInfo(0, column.Index)).SelectMany(_ => _.Colspan().ToOption());
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
            }).Match(_ => _, table.Border);
        }

        private static Func<CellInfo, Option<XColor>> BackgroundColor(this Table table)
        {
            var result = new Dictionary<CellInfo, List<BackgroundTuple>>();
            foreach (var row in table.Rows)
            foreach (var column in table.Columns)
            {
                var cell = row.Cells[column.Index];
                if (cell.BackgroundColor().HasValue)
                    for (var i = 0; i < cell.Rowspan().ToOption().ValueOr(1); i++)
                    for (var j = 0; j < cell.Colspan().ToOption().ValueOr(1); j++)
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
                for (var i = 0; i < row.Cells[column.Index].Rowspan().ToOption().ValueOr(1); i++)
                    if (row.TableHeader())
                        hashSet.Add(row.Index);
            return hashSet;
        }

        private class TablePart
        {
            private double TableY { get; }
            public IEnumerable<int> Rows { get; }
            private int Index { get; }
            public TableInfo TableInfo { get; }
            public bool IsFirst => Index == 0;
            public double Y(Section section, XGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document) => 
                IsFirst ? TableY : section.TopMargin(graphics, tableInfos, mode, document);

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
                    x + paragraph.LeftMargin().ToOption().ValueOr(0),
                    y + paragraph.TopMargin().ToOption().ValueOr(0),
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
                    var lineSpace = font.GetHeight(xGraphics);
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
            var callerInfos = cell.CallerInfos ?? new List<CallerInfo>();
            if (callerInfos.Count <= 0) return;
            var text = string.Join(" ", callerInfos.Select(_ => _.Line));
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