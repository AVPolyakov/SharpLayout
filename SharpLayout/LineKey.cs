using System;

namespace SharpLayout
{
    public struct LineKey : IEquatable<LineKey>
    {
        public int RowIndex { get; }
        public int ColumnIndex { get; }
        public int ElementIndex { get; }
        public int SoftLineIndex { get; }

        public LineKey(int rowIndex, int columnIndex, int elementIndex, int softLineIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            ElementIndex = elementIndex;
            SoftLineIndex = softLineIndex;
        }

        public bool Equals(LineKey other)
        {
            return RowIndex == other.RowIndex && 
                   ColumnIndex == other.ColumnIndex && 
                   ElementIndex == other.ElementIndex && 
                   SoftLineIndex == other.SoftLineIndex;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is LineKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = RowIndex;
                hashCode = (hashCode * 397) ^ ColumnIndex;
                hashCode = (hashCode * 397) ^ ElementIndex;
                hashCode = (hashCode * 397) ^ SoftLineIndex;
                return hashCode;
            }
        }
    }
}