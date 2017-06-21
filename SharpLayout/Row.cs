using System.Collections.Generic;

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

        public Cell this[Column column] => this[column.Index];

        public Cell this[int column] => Cells[column];
    }
}