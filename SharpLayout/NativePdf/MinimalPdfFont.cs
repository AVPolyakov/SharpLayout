using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SharpLayout.NativePdf
{
    public class MinimalPdfFont
    {
        public byte[] RawBytes { get; }
        public TrueTypeMetrics Metrics { get; }
        internal string Alias { get; set; }
        internal HashSet<int> UsedCodePoints { get; } = new HashSet<int>();

        public MinimalPdfFont(byte[] ttfBytes)
        {
            RawBytes = ttfBytes;
            Metrics = new TrueTypeMetrics(ttfBytes);
        }
    }

    public class MinimalPdfPage
    {
        public double Width { get; }
        public double Height { get; }
        private readonly StringBuilder _sb = new StringBuilder();

        public MinimalPdfPage(double width, double height)
        {
            Width = width;
            Height = height;
        }

        public void DrawString(string text, MinimalPdfFont font, double fontSize, double x, double yTop)
        {
            if (string.IsNullOrEmpty(text)) return;

            var m = font.Metrics;
            double baselineY = Height - yTop;

            var hex = new StringBuilder("<");
            for (int i = 0; i < text.Length; i++)
            {
                int cp = char.ConvertToUtf32(text, i);
                if (char.IsSurrogatePair(text, i)) i++;
                font.UsedCodePoints.Add(cp);
                hex.AppendFormat("{0:X4}", m.GetGlyphId(cp));
            }
            hex.Append(">");

            _sb.AppendFormat(CultureInfo.InvariantCulture,
                "BT 0 0 0 rg /{0} {1:0.###} Tf {2:0.###} {3:0.###} Td {4} Tj ET\n",
                font.Alias, fontSize, x, baselineY, hex);
        }

        internal string GetContent() => _sb.ToString();
    }
}