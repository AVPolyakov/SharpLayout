using System;
using PdfSharp.Drawing;

namespace SharpLayout
{
    internal static class ImageRenderer
    {
        public static void Draw(Image image, XUnit x0, XUnit y0, Drawer drawer, HorizontalAlign alignment, double width)
        {
            var y = y0 + image.TopMargin().ToOption().ValueOr(0);
            var x = x0 + image.LeftMargin().ToOption().ValueOr(0);
            var innerWidth = image.GetInnerWidth(width);
            double dx;
            switch (alignment)
            {
	            case HorizontalAlign.Left:
	            case HorizontalAlign.Justify:
		            dx = 0;
		            break;
	            case HorizontalAlign.Center:
		            dx = (innerWidth - image.ContentWidth()) / 2;
		            break;
	            case HorizontalAlign.Right:
		            dx = innerWidth - image.ContentWidth();
		            break;
	            default:
		            throw new ArgumentOutOfRangeException(nameof(alignment), alignment, null);
            }
            drawer.DrawImage(image, x + dx, y);
        }

        private static double ContentWidth(this Image image)
        { 
	        if (image.Width().HasValue)
		        return image.Width().Value;
	        else
	        {
		        var content = image.Content();
		        if (content.HasValue)
			        return content.Value.Process(xImage => xImage.PointWidth);
		        else
			        return 0;
	        }
        }

        public static double GetInnerWidth(this Image image, double width) 
	        => width - image.LeftMargin().ToOption().ValueOr(0) - image.RightMargin().ToOption().ValueOr(0);
    }
}