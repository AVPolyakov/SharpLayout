namespace SharpLayout
{
    public struct ParagraphKey
    {
        public int RowIndex { get; }
        public int ColumnIndex { get; }
        public int ElementIndex { get; }

        public ParagraphKey(int rowIndex, int columnIndex, int elementIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            ElementIndex = elementIndex;
        }
    }
}