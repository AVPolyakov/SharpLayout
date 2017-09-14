using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Table : IElement
    {
        public int Line { get; }

        public readonly List<Column> Columns = new List<Column>();
        
        public readonly List<Row> Rows = new List<Row>();

        public Table([CallerLineNumber] int line = 0)
        {
            Line = line;
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

        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table) => table(this);

        public HorizontalAlign HorizontalAlign { get; set; }

        public Option<double> TopMargin { get; set; }

        public Option<double> BottomMargin { get; set; }

        public Option<double> LeftMargin { get; set; }

        public Option<double> RightMargin { get; set; }

        public Table Margin(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftMargin = value;
            if (direction.HasFlag(Right)) RightMargin = value;
            if (direction.HasFlag(Top)) TopMargin = value;
            if (direction.HasFlag(Bottom)) BottomMargin = value;
            return this;
        }

        public double ColumnsWidth => Columns.Sum(_ => _.Width);
    }
}