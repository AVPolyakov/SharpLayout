using System;
using System.IO;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public interface IGraphics
    {
        void DrawImage(Func<Stream> imageContent, double x, double y);
        void DrawImage(Func<Stream> imageContent, double x, double y, Func<double, double> width, Func<double, double> height);
        void DrawString(string s, Font font, XBrush brush, double x, double y);
        void DrawLine(XPen pen, double x1, double y1, double x2, double y2);
        void DrawRectangle(XBrush brush, double x, double y, double width, double height);
        XSize MeasureString(string text, XFont font, XStringFormat stringFormat);
        void DrawLines(XPen pen, XPoint[] points);
    }

    public class Graphics: IGraphics
    {
        private readonly XGraphics xGraphics;

        public Graphics(XGraphics xGraphics)
        {
            this.xGraphics = xGraphics;
        }

        public void DrawImage(Func<Stream> imageContent, double x, double y)
        {
            using (var stream = imageContent())
            using (var xImage = XImage.FromStream(stream))
                xGraphics.DrawImage(xImage, x, y);
        }

        public void DrawImage(Func<Stream> imageContent, double x, double y, 
            Func<double, double> width, Func<double, double> height)
        {
            using (var stream = imageContent())
            using (var xImage = XImage.FromStream(stream))
                xGraphics.DrawImage(xImage, x, y, width(xImage.PointWidth), height(xImage.PointHeight));
        }

        public void DrawString(string s, Font font, XBrush brush, double x, double y)
        {
            xGraphics.DrawString(s, font.XFont, brush, x, y);
        }

        public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
        {
            xGraphics.DrawLine(pen, x1, y1, x2, y2);
        }

        public void DrawLines(XPen pen, XPoint[] points)
        {
            xGraphics.DrawLines(pen, points);
        }

        public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
        {
            xGraphics.DrawRectangle(brush, x, y, width, height);
        }

        public XSize MeasureString(string text, XFont font, XStringFormat stringFormat)
        {
            return xGraphics.MeasureString(text, font, stringFormat);
        }
    }
}