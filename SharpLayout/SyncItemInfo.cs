using System.Collections.Generic;

namespace SharpLayout
{
    public class SyncBitmapInfo
    {
        public double Resolution { get; set; }
        public int HorizontalPixelCount { get; set; }
        public int VerticalPixelCount { get; set; }
        public SyncPageInfo PageInfo { get; set; }
    }

    public class SyncPageInfo
    {
        public List<SyncItemInfo> ItemInfos = new List<SyncItemInfo>();
    }

    public class SyncItemInfo
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        public int TableLevel { get; set; }
        public List<CallerInfo> CallerInfos { get; set; }
        public int Level { get; set; }
    }

    public class CallerInfo
    {
        public int Line { get; set; }
        public string FilePath { get; set; }
    }
}