using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Cell
    {
        public Table Table { get; }

        public int RowIndex { get; }
        
        public int ColumnIndex { get; }

        private Option<int> rowspan;

        public Cell Rowspan(Option<int> value)
        {
            rowspan = value;
            return this;
        }

        public Option<int> Rowspan() => rowspan;

        private Option<int> colspan;

        public Cell Colspan(int value)
        {
            colspan = value;
            return this;
        }

        public Option<int> Colspan() => colspan;

        public Option<double> LeftBorder { get; set; }

        public Option<double> RightBorder { get; set; }

        public Option<double> TopBorder { get; set; }

        public Option<double> BottomBorder { get; set; }

        public Cell Border(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftBorder = value;
            if (direction.HasFlag(Right)) RightBorder = value;
            if (direction.HasFlag(Top)) TopBorder = value;
            if (direction.HasFlag(Bottom)) BottomBorder = value;
            return this;
        }

        private Option<XColor> backgroundColor;

        public Cell BackgroundColor(XColor value)
        {
            backgroundColor = value;
            return this;
        }

        public Option<XColor> BackgroundColor() => backgroundColor;

        public readonly List<IElement> Elements = new List<IElement>();

        public IEnumerable<Paragraph> Paragraphs => Elements
            .Where(_ => _.Match(p => true, t => false))
            .Select(_ => _.Match(p => p, t => { throw new ApplicationException(); }));

        public Cell Add(IElement value)
        {
            Elements.Add(value);
            return this;
        }

        private VerticalAlign verticalAlign;

        public Cell VerticalAlign(VerticalAlign value)
        {
            verticalAlign = value;
            return this;
        }

        public VerticalAlign VerticalAlign() => verticalAlign;

        public readonly List<CallerInfo> CallerInfos = new List<CallerInfo>();

        internal Cell(Table table, int rowIndex, int columnIndex)
        {
            Table = table;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
        }
    }
}