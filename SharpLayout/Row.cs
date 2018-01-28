using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpLayout
{
    public class Row
    {
        public Table Table { get; }

        public int Index { get; }
        
        internal readonly List<Cell> Cells = new List<Cell>();

        private Option<double> height;

        public Row Height(double value)
        {
            height = value;
            return this;
        }

        public Option<double> Height() => height;

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
        public Row TableHeader(bool value)
        {
            tableHeader = value;
            return this;
        }

        public bool TableHeader() => tableHeader;
    }
}