using System.Collections.Generic;

namespace SharpLayout
{
    public class SyncBitmapInfo
    {
        public int Resolution { get; set; }
        public int HorizontalPixelCount { get; set; }
        public int VerticalPixelCount { get; set; }
        public SyncPageInfo PageInfo { get; set; }
    }

    public class SyncPageInfo
    {
        public List<SyncCellInfo> CellInfos = new List<SyncCellInfo>();
    }

    public class SyncCellInfo
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public int? Line { get; set; }
        public string FilePath { get; set; }
        public int TableLevel { get; set; }
    }
}