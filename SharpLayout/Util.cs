using PdfSharp.Drawing;

namespace SharpLayout
{
    public static class Util
    {
        public static void Colspan(this Cell cell, Column dateColumn) => cell.Colspan = dateColumn.Index - cell.ColumnIndex + 1;

        public static double Px(double value) => XUnit.FromCentimeter(value / 100d);

        public static double Cm(double value) => XUnit.FromCentimeter(value);
    }
}