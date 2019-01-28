using PdfSharp.Drawing;

namespace SharpLayout
{
    internal static class ImageRenderer
    {
        public static void Draw(Image image, XUnit x0, XUnit y0, Drawer drawer)
        {
            var y = y0 + image.TopMargin().ToOption().ValueOr(0);
            var x = x0 + image.LeftMargin().ToOption().ValueOr(0);
            drawer.DrawImage(image, x, y);
        }
    }
}