using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpLayout
{
    public class Cell
    {
        public Table Table { get; }

        public int RowIndex { get; }
        
        public int ColumnIndex { get; }

        public Option<int> Rowspan { get; set; }

        public Option<int> Colspan { get; set; }

        public Option<double> LeftBorder { get; set; }

        public Option<double> RightBorder { get; set; }

        public Option<double> TopBorder { get; set; }

        public Option<double> BottomBorder { get; set; }

        public readonly List<IElement> Elements = new List<IElement>();

        public IEnumerable<Paragraph> Paragraphs => Elements
            .Where(_ => _.Match(p => true, t => false))
            .Select(_ => _.Match(p => p, t => { throw new ApplicationException(); }));

        public void Add(IElement value) => Elements.Add(value);

        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Top;

        public Option<int> MergeRight
        {
            get { return Colspan.Select(_ => _ - 1); }
            set { Colspan = value.Select(_ => _ + 1); }
        }

        public Option<int> MergeDown
        {
            get { return Rowspan.Select(_ => _ - 1); }
            set { Rowspan = value.Select(_ => _ + 1); }
        }

        public readonly List<CallerInfo> CallerInfos = new List<CallerInfo>();

        internal Cell(Table table, int rowIndex, int columnIndex)
        {
            Table = table;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
        }
    }
}