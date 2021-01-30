using PdfSharp.Drawing;

namespace SharpLayout
{
    public static class Dpi254
    {
        public static double Px(double value) => XUnit.FromCentimeter(value / 100d);
    }
}