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

        public Option<Paragraph> Paragraph { get; set; }

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

        public Option<int> Line { get; set; }

        internal Cell(Table table, int rowIndex, int columnIndex)
        {
            Table = table;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
        }
    }
}