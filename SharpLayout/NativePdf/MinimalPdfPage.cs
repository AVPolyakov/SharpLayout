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
        internal Dictionary<ushort, ushort> OldToNewGid { get; set; }

        public MinimalPdfFont(byte[] ttfBytes)
        {
            RawBytes = ttfBytes;
            Metrics = new TrueTypeMetrics(ttfBytes);
        }

        internal ushort GetFinalGlyphId(int codePoint)
        {
            ushort oldGid = Metrics.GetGlyphId(codePoint);
            if (OldToNewGid != null && OldToNewGid.TryGetValue(oldGid, out ushort newGid))
                return newGid;
            return oldGid;
        }
    }

    public class MinimalPdfPage
    {
        public double Width { get; }
        public double Height { get; }

        private readonly List<(string Prefix, MinimalPdfFont Font, string Text, string Suffix)> _chunks =
            new List<(string, MinimalPdfFont, string, string)>();
        private readonly StringBuilder _currentCommands = new StringBuilder();

        public MinimalPdfPage(double width, double height)
        {
            Width = width;
            Height = height;
        }

        public void DrawString(string text, MinimalPdfFont font, double fontSize, Color color, double x, double y)
        {
            if (string.IsNullOrEmpty(text)) return;

            double baselineY = Height - y;

            for (int i = 0; i < text.Length; i++)
            {
                int cp = char.ConvertToUtf32(text, i);
                if (char.IsSurrogatePair(text, i)) i++;
                font.UsedCodePoints.Add(cp);
            }

            _currentCommands.AppendFormat(CultureInfo.InvariantCulture,
                "BT {0:0.###} {1:0.###} {2:0.###} rg /{3} {4:0.####} Tf {5:0.####} {6:0.####} Td <",
                color.R / 255.0, color.G / 255.0, color.B / 255.0,
                font.Alias, fontSize, x, baselineY);

            string prefix = _currentCommands.ToString();
            _currentCommands.Clear();

            _chunks.Add((prefix, font, text, "> Tj ET\n"));
        }

        internal string GetContent()
        {
            var sb = new StringBuilder();
            for (int c = 0; c < _chunks.Count; c++)
            {
                var chunk = _chunks[c];
                sb.Append(chunk.Prefix);

                if (chunk.Font != null && chunk.Text != null)
                {
                    for (int i = 0; i < chunk.Text.Length; i++)
                    {
                        int cp = char.ConvertToUtf32(chunk.Text, i);
                        if (char.IsSurrogatePair(chunk.Text, i)) i++;
                        sb.AppendFormat("{0:X4}", chunk.Font.GetFinalGlyphId(cp));
                    }
                    sb.Append(chunk.Suffix);
                }
            }

            if (_currentCommands.Length > 0)
                sb.Append(_currentCommands);

            return sb.ToString();
        }
    }
}
