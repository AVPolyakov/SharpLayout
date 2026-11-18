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
    }

    public void DrawImage(Func<Stream> imageContent, double x, double y, Func<double, double> width, Func<double, double> height)
    {
    }

    public void DrawString(string s, Font font, XBrush brush, double x, double y)
    {
        _page.DrawString(
            text: s, 
            font: _minimalPdfDocument.GetOrRegisterFont(font),
            fontSize: font.Size,
            x: x,
            yTop: y);
    }

    public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
    {
    }

    public void DrawRectangle(XBrush brush, double x, double y, double width, double height)
    {
    }
    
    public void DrawLines(XPen pen, XPoint[] points)
    {
    }
    
    public XSize MeasureString(string text, Font font, XStringFormat stringFormat)
    {
        var size = FontStorage.GetMetrics(font).MeasureString(text, font.Size);
        return new XSize(size.Width, size.Height);
    }
}