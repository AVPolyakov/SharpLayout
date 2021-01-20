using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using static System.Linq.Enumerable;
using static System.Math;

namespace SharpLayout
{
    public class CharSizeCache
    {
        public readonly Dictionary<int, double> Dictionary = new Dictionary<int, double>(512);
    }

    public class DrawCache
    {
        public readonly Dictionary<FontKey, double> BaseLines = new Dictionary<FontKey, double>();
        private readonly Dictionary<FontKey, CharSizeCache> charSizeCaches = 
            new Dictionary<FontKey, CharSizeCache>();

        public CharSizeCache GetCharSizeCache(Font font)
        {
            var key = new FontKey(font.Name, font.Size, font.Style, font.PdfOptions.FontEncoding);
            if (charSizeCaches.TryGetValue(key, out var value))
                return value;
            var charSizeCache = new CharSizeCache();
            charSizeCaches.Add(key, charSizeCache);
            return charSizeCache;
        }
    }

    internal class RowCache
	{
		public Table Table { get; }
		private readonly Queue<int> queue = new Queue<int>();
		private readonly Dictionary<int, Row> dictionary = new Dictionary<int, Row>();

		public RowCache(Table table)
		{			
			Table = table;
		}

		public Row Row(int rowIndex)
		{
			if (dictionary.TryGetValue(rowIndex, out var value))
				return value;
			var row = Table.RowFuncs[rowIndex]();
			if (queue.Count >= 100)
				dictionary.Remove(queue.Dequeue());
			queue.Enqueue(rowIndex);
			dictionary.Add(rowIndex, row);
			return row;
		}
	}

    internal class PageTuple
    {
        public SyncPageInfo SyncPageInfo { get; }
        public Section Section { get; }

        public PageTuple(SyncPageInfo syncPageInfo, Section section)
        {
            SyncPageInfo = syncPageInfo;
            Section = section;
        }
    }

    public static class TableRenderer
    {
	    internal static RowCache GetRowCache(this Dictionary<Table, RowCache> rowCaches, Table table)
	    {
		    if (rowCaches.TryGetValue(table, out var value))
			    return value;
		    var rowCache = new RowCache(table);
		    rowCaches.Add(table, rowCache);
		    return rowCache;
	    }

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
			    if (!table.KeepWithNext().ToOption().ValueOr(false))
			    {
				    yield return new TableGroup(list);
				    list = new List<Table>();
			    }
		    }
            if (list.Count > 0)
                yield return new TableGroup(list);
	    }

        private class Page
        {
            public List<TablePart> TableParts { get; }
            public List<Table> Footnotes { get; }

            public Page(List<TablePart> tableParts, List<Table> footnotes)
            {
                TableParts = tableParts;
                Footnotes = footnotes;
            }
        }

        private static IEnumerable<Table> GetTables(List<Func<Document, IGraphics, Table[]>> tableFuncList, IGraphics xGraphics,
            Document document)
        {
            return tableFuncList.Select(func => func(document, xGraphics)).SelectMany(_ => _);
        }

        private class PageData
        {
            public Page Page { get; }
            public Section Section { get; }

            public PageData(Page page, Section section)
            {
                Page = page;
                Section = section;
            }
        }

        internal static List<PageTuple> Draw(IGraphics xGraphics, Action<int, Action<IGraphics>, Section> pageAction,
            Document document, GraphicsType graphicsType, SectionGroup sectionGroup, IPageFilter pageFilter)
        {
            var tableInfos = new Dictionary<Table, TableInfo>();
            var rowCaches = new Dictionary<Table, RowCache>();
            var paragraphCaches = new DrawCache();
            var pages = new List<PageData>();
            foreach (var section in sectionGroup.Sections)
            {
                foreach (var tableFuncList in section.tableFuncs)
                {
                    var firstOnPage = true;
                    var y = section.TopMargin(xGraphics, tableInfos, new TextMode.Measure(), document, rowCaches, paragraphCaches);
                    var footnotes = ImmutableQueue.Create<Table>();
                    var tables = GetTables(tableFuncList, xGraphics, document);
                    var tableParts = tables.GetTableGroups().Select(tableGroup => {
                        var slices = SplitByPages(tableGroup, firstOnPage, out var endY, section, y, xGraphics, tableInfos, new TextMode.Measure(), document, rowCaches,
                            footnotes, out var endFootnotes, paragraphCaches);
                        var parts = slices.SelectMany(slice => slice.Rows.Select(
                            (rows, index) => new TablePart(rows, index, slice.TableInfo, slice.TableY))).ToList();
                        if (parts.Count > 0)
                            firstOnPage = false;
                        y = endY;
                        footnotes = endFootnotes;
                        return parts;
                    }).ToList();
                    var result = new List<Page> {new Page(tableParts: new List<TablePart>(), footnotes: new List<Table>())};
                    foreach (var part in tableParts.SelectMany(_ => _))
                    {
                        if (!part.IsFirst)
                            result.Add(new Page(tableParts: new List<TablePart>(), footnotes: new List<Table>()));
                        var page = result[result.Count - 1];
                        page.TableParts.Add(part);
                        var partFootnotes = part.Rows.SelectMany(row => part.TableInfo.Footnotes[row]).ToList();
                        if (partFootnotes.Count > 0)
                        {
                            if (page.Footnotes.Count == 0)
                                page.Footnotes.AddRange(section.FootnoteSeparators);
                            page.Footnotes.AddRange(partFootnotes);
                        }
                    }
                    pages.AddRange(result.Select(page => new PageData(page, section)));
                }
            }
            var syncPageInfos = new List<PageTuple>();
            for (var index = 0; index < pages.Count; index++)
            {
                if (pageFilter.PageMustBeAdd)
                {
                    var section = pages[index].Section;
                    var syncPageInfo = new SyncPageInfo();
                    syncPageInfos.Add(new PageTuple(syncPageInfo, section));

                    void DrawHeaders(Drawer drawer, int pageIndex)
                    {
                        var y0 = 0d;
                        foreach (var header in section.Headers)
                        {
                            Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches).GetValue(header),
                                Range(0, header.RowFuncs.Count), y0: y0, xGraphics, document, tableInfos,
                                section.PageSettings.LeftMargin, syncPageInfo, tableLevel: 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count), rowCaches, section,
                                paragraphCaches);
                            y0 += header.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches);
                        }
                    }

                    void DrawFooters(Drawer drawer, int pageIndex)
                    {
                        var y0 = section.PageSettings.PageHeight -
                            section.Footers.Sum(t => t.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches));
                        foreach (var footer in section.Footers)
                        {
                            Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches).GetValue(footer),
                                Range(0, footer.RowFuncs.Count),
                                y0: y0,
                                xGraphics, document, tableInfos,
                                section.PageSettings.LeftMargin, syncPageInfo, tableLevel: 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count), rowCaches, section,
                                paragraphCaches);
                            y0 += footer.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches);
                        }
                    }

                    void DrawFootnotes(Drawer drawer, int pageIndex)
                    {
                        var y0 = section.PageSettings.PageHeight -
                            section.BottomMargin(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, paragraphCaches) -
                            pages[pageIndex].Page.Footnotes.Sum(t => t.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches));
                        foreach (var footnote in pages[pageIndex].Page.Footnotes)
                        {
                            Draw(GetTableInfo(tableInfos, xGraphics, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches).GetValue(footnote),
                                Range(0, footnote.RowFuncs.Count),
                                y0: y0,
                                xGraphics, document, tableInfos,
                                section.PageSettings.LeftMargin, syncPageInfo, tableLevel: 0, drawer, graphicsType, new TextMode.Draw(pageIndex, pages.Count), rowCaches, section,
                                paragraphCaches);
                            y0 += footnote.GetTableHeight(xGraphics, tableInfos, new TextMode.Draw(pageIndex, pages.Count), document, rowCaches, section, paragraphCaches);
                        }
                    }

                    if (index == 0)
                    {
                        var drawer = new Drawer(xGraphics);
                        DrawHeaders(drawer, index);
                        foreach (var part in pages[index].Page.TableParts)
                            Draw(part.TableInfo, part.Rows, y0: part.Y(section, xGraphics, tableInfos, new TextMode.Draw(index, pages.Count), document, rowCaches, paragraphCaches), xGraphics, document, tableInfos,
                                section.PageSettings.LeftMargin, syncPageInfo, tableLevel: 0, drawer, graphicsType, new TextMode.Draw(index, pages.Count), rowCaches, section,
                                paragraphCaches);
                        DrawFootnotes(drawer, index);
                        DrawFooters(drawer, index);
                        drawer.Flush();
                    }
                    else
                        pageAction(index, xGraphics2 => {
                                var drawer = new Drawer(xGraphics2);
                                DrawHeaders(drawer, index);
                                foreach (var part in pages[index].Page.TableParts)
                                    Draw(part.TableInfo, part.Rows, y0: part.Y(section, xGraphics, tableInfos, new TextMode.Draw(index, pages.Count), document, rowCaches, paragraphCaches),
                                        xGraphics2, document, tableInfos,
                                        section.PageSettings.LeftMargin, syncPageInfo, tableLevel: 0, drawer, graphicsType, new TextMode.Draw(index, pages.Count), rowCaches, section,
                                        paragraphCaches);
                                DrawFootnotes(drawer, index);
                                DrawFooters(drawer, index);
                                drawer.Flush();
                            },
                            section);
                }
                pageFilter.OnNewPage();
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
            IGraphics xGraphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document, Dictionary<Table, RowCache> rowCaches,
            ImmutableQueue<Table> tableFootnotes, out ImmutableQueue<Table> endFootnotes, DrawCache drawCaches)
        {
	        var slices = new List<TableSlice>();
	        var currentTableY = tableY;
	        var currentTableFootnotes = tableFootnotes;
	        var infos = tableGroup.Tables.Select(table => GetTableInfo(tableInfos, xGraphics, new TextMode.Measure(), document, rowCaches, section, drawCaches).GetValue(table)).ToList();
	        var tableGroupFirstPage = true;
	        for (var tableIndex = 0; tableIndex < infos.Count; tableIndex++)
	        {
		        double currentEndY;
	            ImmutableQueue<Table> currentEndFootnotes;
		        if (infos[tableIndex].Table.RowFuncs.Count == 0)
		        {
			        currentEndY = currentTableY + infos[tableIndex].Table.TopMargin().ToOption().ValueOr(0) + 
			                      infos[tableIndex].Table.BottomMargin().ToOption().ValueOr(0);
                    currentEndFootnotes = currentTableFootnotes;
                    slices.Add(new TableSlice(new List<IEnumerable<int>>(), currentTableY, infos[tableIndex]));
		        }
		        else
		        {
			        var mergedRows = infos[tableIndex].MergedRows;
			        var keepWithRows = infos[tableIndex].KeepWithRows;
			        var y = currentTableY + infos[tableIndex].Table.TopMargin().ToOption().ValueOr(0) +
				        infos[tableIndex].Table.Columns.Max(column => infos[tableIndex].TopBorderFunc(new CellInfo(0, column.Index))
					        .Select(_ => _.Width).ValueOr(0));
		            var footnotes = currentTableFootnotes;
                    var lastRowOnPreviousPage = new Option<int>();
			        var row = 0;
			        var tablePieces = new List<IEnumerable<int>>();
			        var addHeader = false;
			        IEnumerable<int> RowRange(int start, int count) =>
				        (!addHeader ? Empty<int>() : infos[tableIndex].TableHeaderRows.OrderBy(_ => _)).Concat(Range(start, count));
		            var bottomMargin = section.BottomMargin(xGraphics, tableInfos, mode, document, rowCaches, drawCaches);
		            var topMargin = section.TopMargin(xGraphics, tableInfos, mode, document, rowCaches, drawCaches);
                    while (true)
		            {
		                y += infos[tableIndex].MaxHeights[row];
		                if (infos[tableIndex].Footnotes[row].Count > 0)
		                    footnotes = (footnotes.IsEmpty ? footnotes.AddRange(section.FootnoteSeparators) : footnotes)
		                        .AddRange(infos[tableIndex].Footnotes[row]);
		                if (section.PageSettings.PageHeight - bottomMargin -
		                    footnotes.Sum(table => table.GetTableHeight(
		                        xGraphics, tableInfos, mode, document, rowCaches, section, drawCaches)) -
		                    y < 0)
		                {
		                    var firstMergedRow = Min(FirstRow(mergedRows, row), FirstRow(keepWithRows, row));
		                    var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
		                    if (firstMergedRow - start > 0)
		                    {
		                        tablePieces.Add(RowRange(start, firstMergedRow - start));
		                        addHeader = true;
		                        lastRowOnPreviousPage = firstMergedRow - 1;
		                        row = firstMergedRow;
		                    }
		                    else
		                    {
		                        if (firstMergedRow == 0 && tableGroupFirstPage && !firstOnPage)
		                        {
		                            tablePieces.Add(Empty<int>());
		                            lastRowOnPreviousPage = new Option<int>();
		                            row = 0;
		                            tableIndex = 0;
		                            slices.Clear();
		                        }
		                        else
		                        {
		                            var endMergedRow = Max(EndRow(infos[tableIndex].Table, mergedRows, row), EndRow(infos[tableIndex].Table, keepWithRows, row));
		                            tablePieces.Add(RowRange(start, endMergedRow - start + 1));
		                            addHeader = true;
		                            lastRowOnPreviousPage = endMergedRow;
		                            row = endMergedRow + 1;
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
		                    y = topMargin + topIndent;
                            footnotes = ImmutableQueue.Create<Table>();
		                }
                        else
		                {
		                    row++;
		                }
		                if (row >= infos[tableIndex].Table.RowFuncs.Count) break;
		            }
                    {
				        var start = lastRowOnPreviousPage.Match(_ => _ + 1, () => 0);
				        if (start < infos[tableIndex].Table.RowFuncs.Count)
					        tablePieces.Add(RowRange(start, infos[tableIndex].Table.RowFuncs.Count - start));
			        }
			        currentEndY = y + infos[tableIndex].Table.BottomMargin().ToOption().ValueOr(0);
                    currentEndFootnotes = footnotes;
                    slices.Add(new TableSlice(tablePieces, currentTableY, infos[tableIndex]));
		        }
		        currentTableY = currentEndY;
                currentTableFootnotes = currentEndFootnotes;
	        }
	        endY = currentTableY;
            endFootnotes = currentTableFootnotes;
	        return slices;
        }

        private static ImmutableQueue<Table> AddRange(this ImmutableQueue<Table> footnotes, List<Table> tables)
        {
            var result = footnotes;
            foreach (var table in tables)
                result = result.Enqueue(table);
            return result;
        }

        private static void Draw(TableInfo info, IEnumerable<int> rows, double y0, IGraphics xGraphics, Document document,
            Dictionary<Table, TableInfo> tableInfos, double x0, SyncPageInfo syncPageInfo, int tableLevel, Drawer drawer, GraphicsType graphicsType,
            TextMode mode, Dictionary<Table, RowCache> rowCaches, Section section,
            DrawCache drawCaches)
        {
            var firstRow = rows.FirstOrNone();
            if (!firstRow.HasValue) return;
            var maxTopBorder = firstRow.Value == 0
                ? MaxTopBorder(info)
                : MaxBottomBorder(firstRow.Value - 1, info.Table, info.BottomBorderFunc);
            var tableY = info.Table.Location().ToOption().Match(_ => _.Y, () => y0) + info.Table.TopMargin().ToOption().ValueOr(0);
            var x0withMargin = info.Table.Location().ToOption().Match(_ => _.X, () => x0) + info.Table.LeftMargin().ToOption().ValueOr(0);
            {
                var x = x0withMargin + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var topBorder = firstRow.Value == 0
                        ? info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index))
                        : info.BottomBorderFunc(new CellInfo(firstRow.Value - 1, column.Index));
                    if (topBorder.HasValue)
                    {
                        var borderY = tableY + maxTopBorder - topBorder.Value.Width/2;
                        var leftBorder = column.Index == 0
                            ? info.LeftBorderFunc(new CellInfo(firstRow.Value, 0)).Select(_ => _.Width)
                            : info.RightBorderFunc(new CellInfo(firstRow.Value, column.Index - 1)).Select(_ => _.Width);
                        drawer.DrawHorizontalLine(
                            topBorder.Value,
                            x,
                            borderY,
                            x + column.Width,
                            leftBorder,
                            bendRight: info.RightBorderFunc(new CellInfo(firstRow.Value, column.Index)).HasValue &&
                            !info.TopBorderFunc(new CellInfo(firstRow.Value, column.Index + 1)).HasValue);
                    }
                    x += column.Width;
                }
            }
            var y = tableY + maxTopBorder;
            var count = rows.Count();
            var index = 0;
            foreach (var row in rows)
            {
                {
                    var leftBorder = info.LeftBorderFunc(new CellInfo(row, 0));
                    if (leftBorder.HasValue)
                    {
                        var borderX = x0withMargin + info.MaxLeftBorder - leftBorder.Value.Width / 2;
                        if ((index == count - 1 ||
                                !info.LeftBorderFunc(new CellInfo(row + 1, 0)).HasValue) &&
                            info.BottomBorderFunc(new CellInfo(row, 0)).HasValue)
                            drawer.DrawVerticalLine(leftBorder.Value, borderX, y,
                                y + info.MaxHeights[row], true);
                        else
                            drawer.DrawLine(leftBorder.Value,
                                borderX, y, borderX, y + info.MaxHeights[row]);
                    }
                }
                var x = x0withMargin + info.MaxLeftBorder;
                foreach (var column in info.Table.Columns)
                {
                    var cell = rowCaches.GetRowCache(info.Table).Row(row).Cells[column.Index];
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
                    HighlightCellLine(xGraphics, info, bottomBorder, row, column, x, y, document, drawer, rowCaches);
                    var cellInnerHeight = Range(0, cell.Rowspan().ToOption().ValueOr(1))
                        .Sum(i => info.MaxHeights[row + i] -
                            MaxBottomBorder(row + cell.Rowspan().ToOption().ValueOr(1) - 1, info.Table, info.BottomBorderFunc));
                    var width = info.Table.ContentWidth(row, column, info.RightBorderFunc, rowCaches);
                    var elements = cell.Elements.Select(_ => new {
                        Element = _,
                        Height = _.Match(
                            p => p.GetParagraphHeight(row, column, info.Table, xGraphics, info.RightBorderFunc, mode, document, rowCaches, drawCaches),
                            t => t.GetTableHeight(xGraphics, tableInfos, mode, document, rowCaches, section, drawCaches),
                            image => image.GetImageHeight())
                    }).ToList();
                    var contentHeight = elements.Sum(_ => _.Height);
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
                    for (var elementIndex = 0; elementIndex < elements.Count; elementIndex++)
                    {
                        var element = elements[elementIndex];
                        if (document.ParagraphsAreHighlighted)
                            element.Element.Match(
                                paragraph => {
                                    HighlightParagraph(paragraph, column, row, x, y + dy + paragraphY, width, info, xGraphics, drawer, mode, document, rowCaches, drawCaches);
                                    return new { };
                                },
                                table => new { },
                                image => new { });
                        element.Element.Match(
                            paragraph => {
                                if (graphicsType == GraphicsType.Image)
                                    syncPageInfo.ItemInfos.Add(new SyncItemInfo {
                                        X = x,
                                        Y = y + dy + paragraphY,
                                        Height = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc, mode, document, rowCaches, drawCaches),
                                        Width = paragraph.GetInnerWidth(width),
                                        CallerInfos = paragraph.CallerInfos ?? new List<CallerInfo>(),
                                        TableLevel = tableLevel,
                                        Level = 1
                                    });
                                ParagraphRenderer.Draw(xGraphics, paragraph, x0: x, y0: y + dy + paragraphY, width, Alignment(paragraph, info.Table), drawer, graphicsType, mode,
                                    document, info.Table, drawCaches);
                                return new { };
                            },
                            table => {
                                var tableInfo = GetTableInfo(tableInfos, xGraphics, mode, document, rowCaches, section, drawCaches).GetValue(table);
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
                                    Range(0, table.RowFuncs.Count), y0: y + dy + paragraphY, xGraphics,
                                    document, tableInfos, x + dx, syncPageInfo, tableLevel + 1, drawer, graphicsType, mode,
                                    rowCaches, section, drawCaches);
                                return new { };
                            },
                            image => {
                                ImageRenderer.Draw(image, x0: x, y0: y + dy + paragraphY, drawer, Alignment(image, info.Table), width);
                                return new { };
                            });
                        paragraphY += element.Height;
                    }
                    var rightBorder = info.RightBorderFunc(new CellInfo(row, column.Index));
                    if (rightBorder.HasValue)
                    {
                        var notExistsRightBorderOnNextRow = index == count - 1 ||
                            !info.RightBorderFunc(new CellInfo(row + 1, column.Index)).HasValue;
                        var borderX = x + column.Width - rightBorder.Value.Width/2;
                        bool? toRight;
                        if (notExistsRightBorderOnNextRow)
                        {
                            if (!info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue &&
                                info.BottomBorderFunc(new CellInfo(row, column.Index)).HasValue)
                                toRight = false;
                            else if (!info.BottomBorderFunc(new CellInfo(row, column.Index)).HasValue &&
                                info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue)
                                toRight = true;
                            else
                                toRight = null;
                        }
                        else
                            toRight = null;
                        if (toRight.HasValue)
                            drawer.DrawVerticalLine(rightBorder.Value, borderX, y,
                                y + info.MaxHeights[row], toRight.Value);
                        else
                        {
                            double bottomBorderWidth;
                            if (notExistsRightBorderOnNextRow &&
                                info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue)
                                bottomBorderWidth = info.BottomBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
                            else
                                bottomBorderWidth = 0d;
                            drawer.DrawLine(rightBorder.Value,
                                borderX, y, borderX, y + info.MaxHeights[row] - bottomBorderWidth);
                        }
                    }
                    if (bottomBorder.HasValue)
                    {
                        Option<double> leftBorder;
                        if (column.Index == 0)
                            leftBorder = !info.LeftBorderFunc(new CellInfo(row, column.Index)).HasValue
                                ? info.LeftBorderFunc(new CellInfo(row + 1, column.Index))
                                    .Select(_ => _.Width)
                                : new Option<double>();
                        else
                        {
                            var leftCell = new CellInfo(row, column.Index - 1);
                            leftBorder = !info.RightBorderFunc(leftCell).HasValue &&
                                         !info.BottomBorderFunc(leftCell).HasValue
                                ? info.RightBorderFunc(new CellInfo(row + 1, column.Index - 1))
                                    .Select(_ => _.Width)
                                : new Option<double>();
                        }
                        double rightBorderWidth;
                        if (!info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue)
                            rightBorderWidth = info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
                        else
                            rightBorderWidth = 0d;
                        var borderY = y + info.MaxHeights[row] - bottomBorder.Value.Width/2;
                        drawer.DrawHorizontalLine(
                            bottomBorder.Value,
                            x,
                            borderY,
                            x + column.Width - rightBorderWidth,
                            leftBorder,
                            bendRight: index != count - 1 &&
                            info.RightBorderFunc(new CellInfo(row + 1, column.Index)).HasValue &&
                            !info.RightBorderFunc(new CellInfo(row, column.Index)).HasValue &&
                            !info.BottomBorderFunc(new CellInfo(row, column.Index + 1)).HasValue);
                    }
                    x += column.Width;
                }
                y += info.MaxHeights[row];
                index++;
            }
        }

        private static void DrawHorizontalLine(this Drawer drawer, XPen pen,
            double x1, double y1, double x2, Option<double> leftBorder, bool bendRight)
        {
            if (leftBorder.HasValue)
            {
                if (bendRight)
                {
                    var newX1 = x1 - leftBorder.Value;
                    var d = pen.Width / 2;
                    drawer.DrawLines(pen, new[] {
                        new XPoint(newX1 + d, y1 + d),
                        new XPoint(newX1 + d, y1),
                        new XPoint(x2 - d, y1),
                        new XPoint(x2 - d, y1 + d),
                    });
                }
                else
                {
                    var newX1 = x1 - leftBorder.Value;
                    var d = pen.Width / 2;
                    drawer.DrawLines(pen, new[] {
                        new XPoint(x2, y1),
                        new XPoint(newX1 + d, y1),
                        new XPoint(newX1 + d, y1 + d),
                    });
                }
            }
            else
            {
                if (bendRight)
                {
                    var d = pen.Width / 2;
                    drawer.DrawLines(pen, new[] {
                        new XPoint(x1, y1),
                        new XPoint(x2 - d, y1),
                        new XPoint(x2 - d, y1 + d),
                    });
                }
                else
                {
                    drawer.DrawLine(pen,
                        x1, y1, x2, y1);
                }
            }
        }

        private static void DrawVerticalLine(this Drawer drawer, XPen pen, double x1, double y1, double y2, bool toRight)
        {
            var d = pen.Width / 2;
            if (toRight)
                drawer.DrawLines(pen, new[] {
                    new XPoint(x1, y1),
                    new XPoint(x1, y2 - d),
                    new XPoint(x1 + d, y2 - d),
                });
            else
                drawer.DrawLines(pen, new[] {
                    new XPoint(x1, y1),
                    new XPoint(x1, y2 - d),
                    new XPoint(x1 - d, y2 - d),
                });
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

	    private static HorizontalAlign Alignment(Image image, Table table)
	    {
		    if (image.Alignment().HasValue) return image.Alignment().Value;
		    if (table.ContentAlign().HasValue) return table.ContentAlign().Value;
			return HorizontalAlign.Left;
	    }

	    private static double TopMargin(this Section section, IGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document,
		    Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches)
        {
            return Max(section.PageSettings.TopMargin,
                section.Headers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document, rowCaches, section, drawCaches)));
        }

        private static double BottomMargin(this Section section, IGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document,
	        Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches)
        {
            return Max(section.PageSettings.BottomMargin,
                section.Footers.Sum(table => table.GetTableHeight(graphics, tableInfos, mode, document, rowCaches, section, drawCaches)));
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
            if (row + 1 >= table.RowFuncs.Count) return row;
            var i = row + 1;
            while (true)
            {
                if (!set.Contains(i))
                    return i - 1;
                i++;
            }
        }

        private static int FirstRow(HashSet<int> set, int row)
        {
            var i = row;
            while (true)
            {
                if (!set.Contains(i))
                    return i;
                i--;
            }
        }

        private static void MergedRows(Row row, Column column, HashSet<int> mergedRows)
        {
            var rowspan = row.Cells[column.Index].Rowspan();
            if (rowspan.HasValue)
                for (var i = row.Index + 1; i < row.Index + rowspan.Value; i++)
                    mergedRows.Add(i);
        }

        private static void KeepWithRows(Row row, HashSet<int> keepWithRows)
        {
            var keepWith = row.KeepWith();
            if (keepWith.HasValue)
                for (var i = row.Index + 1; i < row.Index + keepWith.Value; i++)
                    keepWithRows.Add(i);
        }

        private static double ContentWidth(this Table table, int row, Column column, Func<CellInfo, Option<XPen>> rightBorderFunc,
	        Dictionary<Table, RowCache> rowCaches)
            => table.Find(new CellInfo(row, column.Index), rowCaches).SelectMany(_ => _.Colspan().ToOption()).Match(
                colspan => column.Width
                    + Range(column.Index + 1, colspan - 1).Sum(i => table.Columns[i].Width)
                    - table.BorderWidth(row, column, column.Index + colspan - 1, rightBorderFunc, rowCaches),
                () => column.Width - table.BorderWidth(row, column, column.Index, rightBorderFunc, rowCaches));

        private static double BorderWidth(this Table table, int row, Column column, int rightColumn, Func<CellInfo, Option<XPen>> rightBorderFunc,
	        Dictionary<Table, RowCache> rowCaches)
            => table.Find(new CellInfo(row, column.Index), rowCaches).SelectMany(_ => _.Rowspan().ToOption()).Match(
                rowspan => Range(row, rowspan)
                    .Max(i => rightBorderFunc(new CellInfo(i, rightColumn)).Select(_ => _.Width).ValueOr(0)),
                () => rightBorderFunc(new CellInfo(row, rightColumn)).Select(_ => _.Width).ValueOr(0));

        private static Dictionary<int, double> MaxHeights(this Table table, IGraphics graphics, Func<CellInfo, Option<XPen>> rightBorderFunc,
            Func<CellInfo, Option<XPen>> bottomBorderFunc, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document, Dictionary<Table, RowCache> rowCaches,
            Section section, DrawCache drawCaches)
        {
            var cellContentsByBottomRow = new Dictionary<CellInfo, CellInfo>();
            var heights = new double?[table.RowFuncs.Count];
            {
                var rowIndex = 0;
                foreach (var row in table.Rows)
                {
                    foreach (var column in table.Columns)
                    {
                        var cell = row.Cells[column.Index];
                        var elements = cell.Elements;
                        if (elements.Count > 0)
                        {
                            var rowspan = cell.Rowspan().ToOption();
                            var rowIndex2 = rowspan.Match(value => row.Index + value - 1, () => row.Index);
                            cellContentsByBottomRow.Add(new CellInfo(rowIndex2, column.Index),
                                new CellInfo(row, column));
                        }
                    }
                    heights[rowIndex] = row.Height();
                    rowIndex++;
                }
            }
            var result = new Dictionary<int, double>();
            for (var rowIndex = 0; rowIndex < table.RowFuncs.Count; rowIndex++)
            {
                var maxHeight = 0d;
                foreach (var column in table.Columns)
                {
                    double rowHeightByContent;
                    if (cellContentsByBottomRow.TryGetValue(new CellInfo(rowIndex, column.Index), out var cellInfo))
                    {
						var cell = rowCaches.GetRowCache(table).Row(cellInfo.RowIndex).Cells[cellInfo.ColumnIndex];
                        var paragraphHeight = cell.Elements.Sum(_ => _.Match(
                            p => p.GetParagraphHeight(cell.RowIndex, column, table, graphics, rightBorderFunc, mode, document, rowCaches, drawCaches),
                            t => t.GetTableHeight(graphics, tableInfos, mode, document, rowCaches, section, drawCaches),
                            image => image.GetImageHeight()));
                        rowHeightByContent = cell.Rowspan().ToOption().Match(
                            _ => Max(paragraphHeight - Range(1, _ - 1).Sum(i => result[rowIndex - i]), 0),
                            () => paragraphHeight);
                    }
                    else
                        rowHeightByContent = 0;
                    var innerHeight = heights[rowIndex].ToOption().Match(
                        _ => Max(rowHeightByContent, _), () => rowHeightByContent);
                    var height = innerHeight + MaxBottomBorder(rowIndex, table, bottomBorderFunc);
                    if (maxHeight < height)
                        maxHeight = height;
                }
                result.Add(rowIndex, maxHeight);
            }
            return result;
        }

        private static double GetTableHeight(this Table table, IGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document,
	        Dictionary<Table, RowCache> rowCaches, Section section, DrawCache drawCaches)
        {
            var tableInfo = GetTableInfo(tableInfos, graphics, mode, document, rowCaches, section, drawCaches).GetValue(table);
            return MaxTopBorder(tableInfo) + table.RowFuncs.Select((func, i) => i).Sum(i => tableInfo.MaxHeights[i]) + table.TopMargin().ToOption().ValueOr(0) +
                table.BottomMargin().ToOption().ValueOr(0);
        }

        private static Lazy<Table, TableInfo> GetTableInfo(Dictionary<Table, TableInfo> tableInfos, IGraphics graphics, TextMode mode, Document document,
	        Dictionary<Table, RowCache> rowCaches, Section section, DrawCache drawCaches)
        {
            return Lazy.Create(tableInfos, table => GetTableInfo(graphics, table, tableInfos, mode, document, rowCaches, section, drawCaches));
        }

        private static double GetImageHeight(this Image image)
        {
            double height;
            if (image.Height().HasValue)
                height = image.Height().Value;
            else
            {
                var content = image.Content();
                if (content.HasValue)
                    using (var stream = content.Value())
                    using (var xImage = XImage.FromStream(stream))
                        height = xImage.PointHeight;
                else
                    height = 0;
            }
            return height + image.TopMargin().ToOption().ValueOr(0) + image.BottomMargin().ToOption().ValueOr(0);
        }

        private static double GetParagraphHeight(this Paragraph paragraph, int row, Column column, Table table, IGraphics graphics,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document, Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches)
        {
            return paragraph.GetInnerHeight(graphics, table, row, column, rightBorderFunc, mode, document, rowCaches, drawCaches) +
                paragraph.TopMargin().ToOption().ValueOr(0) + paragraph.BottomMargin().ToOption().ValueOr(0);
        }

        private static double GetInnerHeight(this Paragraph paragraph, IGraphics graphics, Table table, int row, Column column,
            Func<CellInfo, Option<XPen>> rightBorderFunc, TextMode mode, Document document, Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches)
        {
            return ParagraphRenderer.GetHeight(graphics, paragraph, table.ContentWidth(row, column, rightBorderFunc, rowCaches), mode, document, table, drawCaches);
        }

        private static Func<CellInfo, Option<XPen>> RightBorder(this Table table, HashSet<CellInfo> rightMergedCells, 
            Dictionary<CellInfo, List<BorderTuple>> rightBorderDictionary)
        {
            return cell => rightBorderDictionary.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The right border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToString(table)}");
                else
                    return list[0].Value;
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.RowFuncs.Count) return new Option<XPen>();
                if (rightMergedCells.Contains(new CellInfo(cell.RowIndex, cell.ColumnIndex + 1))) return new Option<XPen>();
                return table.Border();
            });
        }

        private static void RightBorderDictionary(Table table, Row row, Column column, Dictionary<CellInfo, List<BorderTuple>> result,
	        Dictionary<Table, RowCache> rowCaches)
        {
            var cell = row.Cells[column.Index];
            var rightBorder = cell.RightBorder();
            if (rightBorder.HasValue)
            {
                var mergeRight = cell.Colspan().ToOption().Match(_ => _ - 1, () => 0);
                result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                    new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
                var rowspan = cell.Rowspan().ToOption();
                if (rowspan.HasValue)
                    for (var i = 1; i <= rowspan.Value - 1; i++)
                        result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                            new BorderTuple(rightBorder.Value, new CellInfo(row, column)));
            }

            var adjacentCell = table.Find(new CellInfo(row.Index, column.Index + 1), rowCaches);
            var leftBorder = adjacentCell.SelectMany(_ => _.LeftBorder());
            if (leftBorder.HasValue)
            {
                var mergeRight = cell.Colspan().ToOption().Match(_ => _ - 1, () => 0);
                result.Add(new CellInfo(row.Index, column.Index + mergeRight),
                    new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
                var rowspan = adjacentCell.SelectMany(_ => _.Rowspan().ToOption());
                if (rowspan.HasValue)
                    for (var i = 1; i <= rowspan.Value - 1; i++)
                        result.Add(new CellInfo(row.Index + i, column.Index + mergeRight),
                            new BorderTuple(leftBorder.Value, new CellInfo(row.Index, column.Index + 1)));
            }
        }

        private static void BottomMergedCells(Row row, Column column, HashSet<CellInfo> bottomMergedCells)
        {
            var cell = row.Cells[column.Index];
            var rowspan = cell.Rowspan().ToOption().ValueOr(1);
            var colspan = cell.Colspan().ToOption().ValueOr(1);
            for (var i = row.Index + 1; i < row.Index + rowspan; i++)
            for (var j = column.Index; j < column.Index + colspan; j++)
                bottomMergedCells.Add(new CellInfo(i, j));
        }

        private static void RightMergedCells(Row row, Column column, HashSet<CellInfo> rightMergedCells)
        {
            var cell = row.Cells[column.Index];
            var rowspan = cell.Rowspan().ToOption().ValueOr(1);
            var colspan = cell.Colspan().ToOption().ValueOr(1);
            for (var i = row.Index; i < row.Index + rowspan; i++)
            for (var j = column.Index + 1; j < column.Index + colspan; j++)
                rightMergedCells.Add(new CellInfo(i, j));
        }

        private static Func<CellInfo, Option<XPen>> BottomBorder(this Table table, HashSet<CellInfo> bottomMergedCells, Dictionary<CellInfo, List<BorderTuple>> bottomBorderDictionary)
        {
            return cell => bottomBorderDictionary.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The bottom border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToString(table)}");
                else
                    return list[0].Value;
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.RowFuncs.Count) return new Option<XPen>();
                if (bottomMergedCells.Contains(new CellInfo(cell.RowIndex + 1, cell.ColumnIndex))) return new Option<XPen>();
                return table.Border();
            });
        }

        private static void BottomBorderDictionary(Table table, Row row, Column column, Dictionary<CellInfo, List<BorderTuple>> result,
	        Dictionary<Table, RowCache> rowCaches)
        {
            var bottomBorder = row.Cells[column.Index].BottomBorder();
            var cell = row.Cells[column.Index];
            if (bottomBorder.HasValue)
            {
                var mergeDown = cell.Rowspan().ToOption().Match(_ => _ - 1, () => 0);
                result.Add(new CellInfo(row.Index + mergeDown, column.Index),
                    new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
                var colspan = cell.Colspan().ToOption();
                if (colspan.HasValue)
                    for (var i = 1; i <= colspan.Value - 1; i++)
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                            new BorderTuple(bottomBorder.Value, new CellInfo(row, column)));
            }

            var adjacentCell = table.Find(new CellInfo(row.Index + 1, column.Index), rowCaches);
            var topBorder = adjacentCell.SelectMany(_ => _.TopBorder());
            if (topBorder.HasValue)
            {
                var mergeDown = cell.Rowspan().ToOption().Match(_ => _ - 1, () => 0);
                result.Add(new CellInfo(row.Index + mergeDown, column.Index),
                    new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
                var colspan = adjacentCell.SelectMany(_ => _.Colspan().ToOption());
                if (colspan.HasValue)
                    for (var i = 1; i <= colspan.Value - 1; i++)
                        result.Add(new CellInfo(row.Index + mergeDown, column.Index + i),
                            new BorderTuple(topBorder.Value, new CellInfo(row.Index + 1, column.Index)));
            }
        }

        private static Func<CellInfo, Option<XPen>> LeftBorder(this Table table, Dictionary<CellInfo, List<BorderTuple>> leftBorderDictionary)
        {
            return cell => leftBorderDictionary.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The left border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToString(table)}");
                else
                    return list[0].Value;
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.RowFuncs.Count) return new Option<XPen>();
                return table.Border();
            });
        }

        private static void LeftBorderDictionary(Row row, Dictionary<CellInfo, List<BorderTuple>> leftBorderDictionary)
        {
            var cell = row.Cells[0];
            var leftBorder = cell.LeftBorder();
            if (leftBorder.HasValue)
            {
                leftBorderDictionary.Add(new CellInfo(row.Index, 0),
                    new BorderTuple(leftBorder.Value, new CellInfo(row.Index, 0)));
                var rowspan = cell.Rowspan().ToOption();
                if (rowspan.HasValue)
                    for (var i = 1; i <= rowspan.Value - 1; i++)
                        leftBorderDictionary.Add(new CellInfo(row.Index + i, 0),
                            new BorderTuple(leftBorder.Value, new CellInfo(row.Index, 0)));
            }
        }

        private static Func<CellInfo, Option<XPen>> TopBorder(this Table table, Dictionary<Table, RowCache> rowCaches)
        {
            var result = new Dictionary<CellInfo, List<BorderTuple>>();
            foreach (var column in table.Columns)
            {
                var cell = table.Find(new CellInfo(0, column.Index), rowCaches);
                var bottomBorder = cell.SelectMany(_ => _.TopBorder());
                if (bottomBorder.HasValue)
                {
                    result.Add(new CellInfo(0, column.Index),
                        new BorderTuple(bottomBorder.Value, new CellInfo(0, column.Index)));
                    var colspan = cell.SelectMany(_ => _.Colspan().ToOption());
                    if (colspan.HasValue)
                        for (var i = 1; i <= colspan.Value - 1; i++)
                            result.Add(new CellInfo(0, column.Index + i),
                                new BorderTuple(bottomBorder.Value, new CellInfo(0, column.Index)));
                }
            }
            return cell => result.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The top border is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToString(table)}");
                else
                    return list[0].Value;
            }).Match(_ => _, () => {
                if (cell.ColumnIndex < 0 || cell.ColumnIndex >= table.Columns.Count) return new Option<XPen>();
                if (cell.RowIndex < 0 || cell.RowIndex >= table.RowFuncs.Count) return new Option<XPen>();
                return table.Border();
            });
        }

        private static Func<CellInfo, Option<XColor>> BackgroundColor(this Table table, Dictionary<CellInfo, List<BackgroundTuple>> backgroundColorDictionary)
        {
            return cell => backgroundColorDictionary.Get(cell).Select(list => {
                if (list.Count > 1)
                    throw new Exception($"The background color is ambiguous Cells={list.Select(_ => _.CellInfo).CellsToString(table)}");
                else
                    return list[0].Color;
            });
        }

        private static void BackgroundColor(Row row, Column column, Dictionary<CellInfo, List<BackgroundTuple>> backgroundColorDictionary)
        {
            var cell = row.Cells[column.Index];
            if (cell.BackgroundColor().HasValue)
                for (var i = 0; i < cell.Rowspan().ToOption().ValueOr(1); i++)
                for (var j = 0; j < cell.Colspan().ToOption().ValueOr(1); j++)
                    backgroundColorDictionary.Add(new CellInfo(row.Index + i, column.Index + j),
                        new BackgroundTuple(cell.BackgroundColor().Value, new CellInfo(cell)));
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

        private static string CellsToString(this IEnumerable<CellInfo> cells, Table table) 
            => string.Join(", ", cells.Select(_ => $"L{table.Line} r{_.RowIndex + 1}c{_.ColumnIndex + 1}"));

        private static TableInfo GetTableInfo(IGraphics xGraphics, Table table, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document,
	        Dictionary<Table, RowCache> rowCaches, Section section, DrawCache drawCaches)
        {
            var rightMergedCells = new HashSet<CellInfo>();
            var bottomMergedCells = new HashSet<CellInfo>();
            var rightBorderDictionary = new Dictionary<CellInfo, List<BorderTuple>>();
            var bottomBorderDictionary = new Dictionary<CellInfo, List<BorderTuple>>();
            var backgroundColorDictionary = new Dictionary<CellInfo, List<BackgroundTuple>>();
            var keepWithRows = new HashSet<int>();
            var mergedRows = new HashSet<int>();
            var leftBorderDictionary = new Dictionary<CellInfo, List<BorderTuple>>();
            var tableHeaderRows = new HashSet<int>();
            var footnotes = new List<Table>[table.RowFuncs.Count];
            foreach (var row in table.Rows)
            {
                KeepWithRows(row, keepWithRows);
                LeftBorderDictionary(row, leftBorderDictionary);
                foreach (var column in table.Columns)
                {
                    RightMergedCells(row, column, rightMergedCells);
                    MergedRows(row, column, mergedRows);
                    BottomMergedCells(row, column, bottomMergedCells);
                    RightBorderDictionary(table, row, column, rightBorderDictionary, rowCaches);
                    BottomBorderDictionary(table, row, column, bottomBorderDictionary, rowCaches);
                    BackgroundColor(row, column, backgroundColorDictionary);
                    TableHeaderRows(row, column, tableHeaderRows);
                }
                footnotes[row.Index] = row.Cells.SelectMany(cell => cell.Elements
                    .SelectMany(element => element.Match(
                        paragraph => paragraph.Spans.SelectMany(span => span.Footnotes.Select(func => func(section))),
                        table1 => {
                            var tableInfo = GetTableInfo(tableInfos, xGraphics, mode, document, rowCaches, section, drawCaches).GetValue(table1);
                            return table1.RowFuncs.SelectMany((func, i) => tableInfo.Footnotes[i]);
                        },
                        image => new Table[] { }
                    ))).ToList();
            }
            var rightBorderFunc = table.RightBorder(rightMergedCells, rightBorderDictionary);
            var bottomBorderFunc = table.BottomBorder(bottomMergedCells, bottomBorderDictionary);
            return new TableInfo(table, table.TopBorder(rowCaches), bottomBorderFunc,
                table.MaxHeights(xGraphics, rightBorderFunc, bottomBorderFunc, tableInfos, mode, document, rowCaches, section, drawCaches), table.LeftBorder(leftBorderDictionary),
                rightBorderFunc, table.BackgroundColor(backgroundColorDictionary), tableHeaderRows,
                keepWithRows, mergedRows, footnotes);
        }

        private static void TableHeaderRows(Row row, Column column, HashSet<int> tableHeaderRows)
        {
            for (var i = 0; i < row.Cells[column.Index].Rowspan().ToOption().ValueOr(1); i++)
                if (row.TableHeader())
                    tableHeaderRows.Add(row.Index);
        }

        private class TablePart
        {
            public IEnumerable<int> Rows { get; }
            private double TableY { get; }
            private int Index { get; }
            public TableInfo TableInfo { get; }
            public bool IsFirst => Index == 0;
            public double Y(Section section, IGraphics graphics, Dictionary<Table, TableInfo> tableInfos, TextMode mode, Document document,
	            Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches) => 
                IsFirst ? TableY : section.TopMargin(graphics, tableInfos, mode, document, rowCaches, drawCaches);

            public TablePart(IEnumerable<int> rows, int index, TableInfo tableInfo, double tableY)
            {
                Rows = rows;
                TableY = tableY;
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
            public HashSet<int> KeepWithRows { get; }
            public HashSet<int> MergedRows { get; }
            public List<Table>[] Footnotes { get; }
            public Func<CellInfo, Option<XPen>> LeftBorderFunc { get; }
            public Func<CellInfo, Option<XPen>> TopBorderFunc { get; }
            public Func<CellInfo, Option<XPen>> BottomBorderFunc { get; }
            public Dictionary<int, double> MaxHeights { get; }
            public double MaxLeftBorder { get; }

            public TableInfo(Table table, Func<CellInfo, Option<XPen>> topBorderFunc,
                Func<CellInfo, Option<XPen>> bottomBorderFunc,
                Dictionary<int, double> maxHeights, Func<CellInfo, Option<XPen>> leftBorderFunc,
                Func<CellInfo, Option<XPen>> rightBorderFunc,
                Func<CellInfo, Option<XColor>> backgroundColor, HashSet<int> tableHeaderRows,
                HashSet<int> keepWithRows,
                HashSet<int> mergedRows, List<Table>[] footnotes)
            {
                Table = table;
                RightBorderFunc = rightBorderFunc;
                BackgroundColor = backgroundColor;
                TableHeaderRows = tableHeaderRows;
                KeepWithRows = keepWithRows;
                MergedRows = mergedRows;
                Footnotes = footnotes;
                LeftBorderFunc = leftBorderFunc;
                TopBorderFunc = topBorderFunc;
                BottomBorderFunc = bottomBorderFunc;
                MaxHeights = maxHeights;
                MaxLeftBorder = table.RowFuncs.Count == 0
                    ? 0
                    : table.RowFuncs.Select((func, i) => i).Max(i => leftBorderFunc(new CellInfo(i, 0))
                        .Select(_ => _.Width).ValueOr(0));
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

        private static void HighlightParagraph(Paragraph paragraph, Column column, int row, double x, double y, double width, TableInfo info, IGraphics xGraphics, Drawer drawer,
            TextMode mode, Document document, Dictionary<Table, RowCache> rowCaches, DrawCache drawCaches)
        {
            var innerHeight = paragraph.GetInnerHeight(xGraphics, info.Table, row, column, info.RightBorderFunc, mode, document, rowCaches, drawCaches);
            var innerWidth = paragraph.GetInnerWidth(width);
            if (innerWidth > 0 && innerHeight > 0)
                FillRectangle(drawer, XColor.FromArgb(32, 0, 0, 255),
                    x + paragraph.LeftMargin().ToOption().ValueOr(0),
                    y + paragraph.TopMargin().ToOption().ValueOr(0),
                    innerWidth, innerHeight);
        }

        private static void HighlightCells(IGraphics xGraphics, TableInfo info, Option<XPen> bottomBorder, int row, Column column, double x, double y,
            double tableY, Drawer drawer, Document document)
        {
            if (document.CellsAreHighlighted && info.Table.Highlighted())
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
                var font = new Font("Times New Roman", 10, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode));
                var redBrush = new XSolidBrush(XColor.FromArgb(200, 255, 0, 0));
                var purpleBrush = new XSolidBrush(XColor.FromArgb(200, 87, 0, 127));
                if (column.Index == 0)
                {
                    var text = $"r{row + 1}";
                    var lineSpace = font.GetHeight();
                    var rnHeight = lineSpace * font.FontFamily.GetCellAscent(font.Style) / font.FontFamily.GetLineSpacing(font.Style);
                    var rnX = x - xGraphics.MeasureString(text, font.XFont, ParagraphRenderer.MeasureTrailingSpacesStringFormat).Width;
                    drawer.DrawString(text, font, redBrush, rnX, y + rnHeight);
                    if (row == 0)
                    {
                        var lineText = $"t{info.Table.Line}";
                        drawer.DrawString(lineText,
                            font,
                            purpleBrush,
                            rnX - xGraphics.MeasureString($"{lineText} ", font.XFont, ParagraphRenderer.MeasureTrailingSpacesStringFormat).Width,
                            y + rnHeight);
                    }
                }
                if (row == 0)
                    drawer.DrawString($"c{column.Index + 1}", font, redBrush, x, tableY - 1);
            }
        }

        private static void HighlightCellLine(IGraphics xGraphics, TableInfo info, Option<XPen> bottomBorder, int row, Column column, double x, double y, Document document,
            Drawer drawer, Dictionary<Table, RowCache> rowCaches)
        {
            if (!document.CellLineNumbersAreVisible) return;
            var cell = rowCaches.GetRowCache(info.Table).Row(row).Cells[column.Index];
            var callerInfos = cell.CallerInfos ?? new List<CallerInfo>();
            if (callerInfos.Count <= 0) return;
            var text = string.Join(" ", callerInfos.Select(_ => _.Line));
            var font = new Font("Arial", 7, XFontStyle.Regular, new XPdfFontOptions(PdfFontEncoding.Unicode));
            var height = info.MaxHeights[row] - bottomBorder.Select(_ => _.Width).ValueOr(0);
            var width = column.Width - info.RightBorderFunc(new CellInfo(row, column.Index)).Select(_ => _.Width).ValueOr(0);
            drawer.DrawString(text,
                font,
                new XSolidBrush(XColor.FromArgb(128, 0, 0, 255)),
                x + width - xGraphics.MeasureString(text, font.XFont, ParagraphRenderer.MeasureTrailingSpacesStringFormat).Width,
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