using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Span
    {
        public string Text { get; }

        public string TextOrEmpty => Text ?? "";

        public XFont Font { get; }
        
        public XBrush Brush { get; set; } = XBrushes.Black;

        public Span(string text, XFont font)
        {
            Text = text;
            Font = font;
        }
    }
}