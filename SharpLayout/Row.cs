using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpLayout
{
    public class Row
    {
        public Table Table { get; }

        public int Index { get; }
        
        internal readonly List<Cell> Cells = new List<Cell>();

        private double? height;
        public double? Height() => height;
        public Row Height(double? value)
        {
            height = value;
            return this;
        }

        internal Row(Table table, int index)
        {
            Table = table;
            Index = index;
        }

        public Cell this[Column column, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = ""] => 
            this[column.Index, line, filePath];

        public Cell this[int columnIndex, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = ""]
        {
            get
            {
                var cell = Cells[columnIndex];
                cell.CallerInfos.Add(new CallerInfo {Line = line, FilePath = filePath});
                return cell;
            }
        }

        private bool tableHeader;
	    public bool TableHeader() => tableHeader;
        public Row TableHeader(bool value)
        {
            tableHeader = value;
            return this;
        }

	    private int? keepWith;
	    public int? KeepWith() => keepWith;
	    public Row KeepWith(int? value)
	    {
		    keepWith = value;
		    return this;
	    }
    }
}