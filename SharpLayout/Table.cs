using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpLayout
{
    public class Table
    {
        public int Line { get; }
        internal readonly double X0;

        public readonly List<Column> Columns = new List<Column>();
        
        public readonly List<Row> Rows = new List<Row>();

        public Table(double x0,  [CallerLineNumber] int line = 0)
        {
            Line = line;
            X0 = x0;
        }

        public Column AddColumn()
        {
            var column = new Column(Columns.Count);
            Columns.Add(column);
            foreach (var row in Rows)
                row.Cells.Add(new Cell(this, row.Index, column.Index));
            return column;
        }

        public Column AddColumn(double width)
        {
            var column = AddColumn();
            column.Width = width;
            return column;
        }

        public Row AddRow()
        {
            var row = new Row(this, Rows.Count);
            Rows.Add(row);
            foreach (var column in Columns)
                row.Cells.Add(new Cell(this, row.Index, column.Index));
            return row;
        }

        internal Option<Cell> Find(CellInfo cell)
        {
            if (cell.RowIndex >= Rows.Count) return new Option<Cell>();
            if (cell.ColumnIndex >= Columns.Count) return new Option<Cell>();
            return Rows[cell.RowIndex].Cells[cell.ColumnIndex];
        }
    }
}