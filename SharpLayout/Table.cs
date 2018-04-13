using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Table : IElement
    {
        internal bool KeepWithNext { get; }

        public int Line { get; }

        public readonly List<Column> Columns = new List<Column>();
        
        public readonly List<Row> Rows = new List<Row>();

        public Table([CallerLineNumber] int line = 0)
            : this(false, line)
        {
        }

        internal Table(bool keepWithNext, [CallerLineNumber] int line = 0)
        {
            KeepWithNext = keepWithNext;
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

        private HorizontalAlign horizontalAlign;
        public HorizontalAlign HorizontalAlign() => horizontalAlign;
        public Table HorizontalAlign(HorizontalAlign value)
        {
            horizontalAlign = value;
            return this;
        }

        private Option<double> topMargin;
        public Option<double> TopMargin() => topMargin;
        public Table TopMargin(Option<double> value)
        {
            topMargin = value;
            return this;
        }

        private Option<double> bottomMargin;
        public Option<double> BottomMargin() => bottomMargin;
        public Table BottomMargin(Option<double> value)
        {
            bottomMargin = value;
            return this;
        }

        private Option<double> leftMargin;
        public Option<double> LeftMargin() => leftMargin;
        public Table LeftMargin(Option<double> value)
        {
            leftMargin = value;
            return this;
        }

        private Option<double> rightMargin;
        public Option<double> RightMargin() => rightMargin;
        public Table RightMargin(Option<double> value)
        {
            rightMargin = value;
            return this;
        }

        public Table Margin(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftMargin(value);
            if (direction.HasFlag(Right)) RightMargin(value);
            if (direction.HasFlag(Top)) TopMargin(value);
            if (direction.HasFlag(Bottom)) BottomMargin(value);
            return this;
        }

        public double ColumnsWidth => Columns.Sum(_ => _.Width);

        private Option<XPen> border;
        public Option<XPen> Border() => border;
        public Table Border(Option<XPen> value)
        {
            border = value;
            return this;
        }

        public Table Border(double value) => Border(new XPen(XColors.Black, value));
    }
}