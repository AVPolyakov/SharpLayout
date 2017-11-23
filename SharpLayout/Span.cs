using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Span
    {
        public string Text { get; }

        public string TextOrEmpty => Text ?? "";

        public XFont Font { get; }
        
        public XBrush Brush { get; set; } = XBrushes.Black;

        private Option<XColor> backgroundColor;

        public Span BackgroundColor(XColor value)
        {
            backgroundColor = value;
            return this;
        }

        public Option<XColor> BackgroundColor() => backgroundColor;

        public Span(string text, XFont font)
        {
            Text = text;
            Font = font;
        }
    }
}