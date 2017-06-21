using PdfSharp.Drawing;

namespace SharpLayout
{
    public class PageSettings
    {
        public double PageWidth { get; set; } = XUnit.FromMillimeter(210);

        public double PageHeight { get; set; } = XUnit.FromMillimeter(297);

        public double LeftMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double RightMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double TopMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double BottomMargin { get; set; } = XUnit.FromCentimeter(1);

        public bool IsHighlightCells { get; set; } = false;
    }
}