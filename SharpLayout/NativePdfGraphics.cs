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
        throw new NotImplementedException();
    }

    public void DrawImage(Func<Stream> imageContent, double x, double y, Func<double, double> width, Func<double, double> height)
    {
        throw new NotImplementedException();
    }

    public void DrawString(string s, Font font, XBrush brush, double x, double y)
    {
        var color = brush is XSolidBrush solid
            ? new Color(solid.Color.R, solid.Color.G, solid.Color.B)
            : Colors.Black;
        
        _page.DrawString(
            text: s, 
            font: _minimalPdfDocument.GetOrRegisterFont(font),
            fontSize: font.Size,
            color: color,
            x: x,
            y: y);
    }

    public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
    {
        _page.DrawLine(ToColor(pen.Color), pen.Width, x1, y1, x2, y2);   
    }
    
    public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
    {
        if (brush is XSolidBrush solid)
        {
            _page.DrawRectangle(ToColor(solid.Color), null, 0, x, y, width, height);
        }
    }

    public void DrawRectangle(XPen pen, XBrush brush, double x, double y, double width, double height)
    {
        Color? fill = brush is XSolidBrush solid ? ToColor(solid.Color) : (Color?)null;
        Color? stroke = pen != null ? ToColor(pen.Color) : (Color?)null;
        double strokeWidth = pen?.Width ?? 0;

        _page.DrawRectangle(fill, stroke, strokeWidth, x, y, width, height);
    }
    
    public void DrawLines(XPen pen, XPoint[] points)
    {
        var rawPoints = new (double X, double Y)[points.Length];
        for (int i = 0; i < points.Length; i++)
        {
            rawPoints[i] = (points[i].X, points[i].Y);
        }

        _page.DrawLines(ToColor(pen.Color), pen.Width, rawPoints);     
    }
    
    private static Color ToColor(XColor xColor)
        => new (xColor.R, xColor.G, xColor.B, (byte)(xColor.A * 255));
    
    public XSize MeasureString(string text, Font font, XStringFormat stringFormat)
    {
        var size = FontStorage.GetMetrics(font).MeasureString(text, font.Size);
        return new XSize(size.Width, size.Height);
    }
}