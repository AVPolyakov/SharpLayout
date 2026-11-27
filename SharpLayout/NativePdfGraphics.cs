using System;
using System.IO;
using PdfSharp.Drawing;
using SharpLayout.NativePdf;

namespace SharpLayout;

public class NativePdfGraphics : IGraphics
{
    private readonly MinimalPdfPage _page;
    private readonly MinimalPdfDocument _minimalPdfDocument;

    public NativePdfGraphics(MinimalPdfPage page, MinimalPdfDocument minimalPdfDocument)
    {
        _page = page;
        _minimalPdfDocument = minimalPdfDocument;
    }

    public void DrawImage(Func<Stream> imageContent, double x, double y)
    {
        var img = _minimalPdfDocument.GetOrRegisterImage(imageContent);
        _page.DrawImage(img, x, y, img.PointWidth, img.PointHeight);
    }

    public void DrawImage(Func<Stream> imageContent, double x, double y,
        Func<double, double> width, Func<double, double> height)
    {
        var img = _minimalPdfDocument.GetOrRegisterImage(imageContent);
        double w = width != null ? width(img.PointWidth) : img.PointWidth;
        double h = height != null ? height(img.PointHeight) : img.PointHeight;
        _page.DrawImage(img, x, y, w, h);
    }

    public void DrawString(string s, Font font, XBrush brush, double x, double y)
    {
        var color = brush is XSolidBrush solid
            ? new Color(solid.Color.R, solid.Color.G, solid.Color.B)
            : Colors.Black;

        var minimalPdfFont = _minimalPdfDocument.GetOrRegisterFont(font);

        _page.DrawString(
            text: s,
            font: minimalPdfFont,
            fontSize: font.Size,
            color: color,
            x: x,
            y: y);

        if ((font.Underline || font.Strikeout) && !string.IsNullOrEmpty(s))
        {
            var metrics = minimalPdfFont.Metrics;
            double width = metrics.MeasureString(s, font.Size).Width;
            double scale = font.Size / metrics.UnitsPerEm;

            if (font.Underline)
            {
                double underlinePosition = metrics.UnderlinePosition * scale;
                double underlineThickness = metrics.UnderlineThickness * scale;
                _page.DrawRectangle(color, x, y - underlinePosition, width, underlineThickness);
            }

            if (font.Strikeout)
            {
                double strikeoutPosition = metrics.StrikeoutPosition * scale;
                double strikeoutSize = metrics.StrikeoutSize * scale;
                _page.DrawRectangle(color, x, y - strikeoutPosition, width, strikeoutSize);
            }
        }
    }

    public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
    {
        _page.DrawLine(ToColor(pen.Color), pen.Width, x1, y1, x2, y2, GetDashPattern(pen));   
    }
    
    public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
    {
        if (brush is XSolidBrush solid)
        {
            _page.DrawRectangle(ToColor(solid.Color), x, y, width, height);
        }
    }

    public void DrawLines(XPen pen, XPoint[] points)
    {
        var rawPoints = new (double X, double Y)[points.Length];
        for (int i = 0; i < points.Length; i++)
        {
            rawPoints[i] = (points[i].X, points[i].Y);
        }

        _page.DrawLines(ToColor(pen.Color), pen.Width, rawPoints, GetDashPattern(pen));
    }
    
    private static readonly double[] _dashPattern = { 3.0, 1.0 };
    private static readonly double[] _dotPattern = { 1.0, 1.0 };
    private static readonly double[] _dashDotPattern = { 3.0, 1.0, 1.0, 1.0 };
    private static readonly double[] _dashDotDotPattern = { 3.0, 1.0, 1.0, 1.0, 1.0, 1.0 };

    private static double[] GetDashPattern(XPen pen)
    {
        return pen.DashStyle switch
        {
            XDashStyle.Dash => _dashPattern,
            XDashStyle.Dot => _dotPattern,
            XDashStyle.DashDot => _dashDotPattern,
            XDashStyle.DashDotDot => _dashDotDotPattern,
            XDashStyle.Custom => pen.DashPattern,
            _ => null
        };
    }
    
    private static Color ToColor(XColor xColor)
        => new Color(xColor.R, xColor.G, xColor.B, (byte)Math.Round(xColor.A * 255));
    
    public XSize MeasureString(string text, Font font, XStringFormat stringFormat)
    {
        var size = FontStorage.GetMetrics(font).MeasureString(text, font.Size);
        return new XSize(size.Width, size.Height);
    }
}