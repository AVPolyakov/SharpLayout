using System;

namespace SharpLayout
{
    public struct CellInfo: IEquatable<CellInfo>
    {
        public int RowIndex { get; }

        public int ColumnIndex { get; }

        public CellInfo(int rowIndex, int columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        public CellInfo(Row row, Column column) : this(row.Index, column.Index)
        {
        }

        public CellInfo(Cell cell) : this(cell.RowIndex, cell.ColumnIndex)
        {
        }

        public static implicit operator CellInfo(Cell cell) => new CellInfo(cell);

        public bool Equals(CellInfo other) => Tuple.Equals(other.Tuple);

        public override int GetHashCode() => Tuple.GetHashCode();

        public override bool Equals(object obj)
        {
            if (obj is CellInfo) return Equals((CellInfo) obj);
            return false;
        }

        private Tuple<int, int> Tuple => System.Tuple.Create(RowIndex, ColumnIndex);
    }
}