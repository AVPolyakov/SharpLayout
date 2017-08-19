using System;
using PdfSharp;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class PageSettings
    {
        public double PageWidth
        {
            get
            {
                switch (Orientation)
                {
                    case PageOrientation.Portrait:
                        return XUnit.FromMillimeter(210);
                    case PageOrientation.Landscape:
                        return XUnit.FromMillimeter(297);
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        public double PageHeight
        {
            get
            {
                switch (Orientation)
                {
                    case PageOrientation.Portrait:
                        return XUnit.FromMillimeter(297);
                    case PageOrientation.Landscape:
                        return XUnit.FromMillimeter(210);
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        public double LeftMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double RightMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double TopMargin { get; set; } = XUnit.FromCentimeter(1);
        
        public double BottomMargin { get; set; } = XUnit.FromCentimeter(1);

        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

        public bool IsHighlightCells { get; set; } = false;

        public bool IsHighlightCellLine { get; set; } = false;
    }
}