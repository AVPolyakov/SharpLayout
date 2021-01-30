using PdfSharp.Drawing;

namespace SharpLayout
{
    public static class Util
    {
        public static Cell Colspan(this Cell cell, Column column)
        {
            cell.Colspan(column.Index - cell.ColumnIndex + 1);
            return cell;
        }

        public static double Cm(double value) => XUnit.FromCentimeter(value);
        
        public static double Mm(double value) => XUnit.FromMillimeter(value);
    }
}