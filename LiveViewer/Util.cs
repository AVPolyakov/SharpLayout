using System;
using SharpLayout;

namespace LiveViewer
{
    public static class Util
    {
        public static double ToPixel(this double value, SyncBitmapInfo bitmapInfo)
        {
            return Math.Round(value * bitmapInfo.Resolution / 72, MidpointRounding.AwayFromZero);
        }
    }
}