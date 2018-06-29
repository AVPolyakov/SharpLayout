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

        public bool Equals(CellInfo other)
        {
            return RowIndex == other.RowIndex && ColumnIndex == other.ColumnIndex;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is CellInfo info && Equals(info);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (RowIndex * 397) ^ ColumnIndex;
            }
        }
    }
}