using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpLayout
{
    public class Row
    {
        public Table Table { get; }

        public int Index { get; }
        
        internal readonly List<Cell> Cells = new List<Cell>();

        public Option<double> Height { get; set; }

        internal Row(Table table, int index)
        {
            Table = table;
            Index = index;
        }

        public Cell this[Column column, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = ""]
        {
            get
            {
                var cell = Cells[column.Index];
                cell.CallerInfos.Add(new CallerInfo {Line = line, FilePath = filePath});
                return cell;
            }
        }
    }
}