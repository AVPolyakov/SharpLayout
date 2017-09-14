using System.Linq;
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

        public static double Px(double value) => XUnit.FromCentimeter(value / 100d);

        public static double Cm(double value) => XUnit.FromCentimeter(value);
    }
}