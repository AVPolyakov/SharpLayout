using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SharpLayout.NativePdf
{
    public class MinimalPdfPage
    {
        public double Width { get; }
        public double Height { get; }

        private readonly List<(string Prefix, MinimalPdfFont Font, string Text, string Suffix)> _chunks =
            new List<(string, MinimalPdfFont, string, string)>();
        private readonly StringBuilder _currentCommands = new StringBuilder();

        // Состояние для батчинга линий одного стиля
        private bool _hasOpenPath;
        private double _currentStrokeWidth = -1;
        private Color _currentStrokeColor;

        public MinimalPdfPage(double width, double height)
        {
            Width = width;
            Height = height;
        }

        private void FlushStrokePath()
        {
            if (_hasOpenPath)
            {
                _currentCommands.Append("S\n");
                _hasOpenPath = false;
            }
        }

        private void EnsureStrokeStyle(Color color, double lineWidth)
        {
            bool styleChanged = !_hasOpenPath
                || Math.Abs(_currentStrokeWidth - lineWidth) > 0.0001
                || _currentStrokeColor != color;

            if (styleChanged)
            {
                FlushStrokePath();
                _currentStrokeWidth = lineWidth;
                _currentStrokeColor = color;

                _currentCommands.AppendFormat(CultureInfo.InvariantCulture,
                    "{0:0.####} w {1:0.###} {2:0.###} {3:0.###} RG\n",
                    lineWidth, color.R / 255.0, color.G / 255.0, color.B / 255.0);

                _hasOpenPath = true;
            }
        }

        public void DrawLine(Color color, double lineWidth, double x1, double y1, double x2, double y2)
        {
            if (color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth);
            _currentCommands.AppendFormat(CultureInfo.InvariantCulture,
                "{0:0.####} {1:0.####} m {2:0.####} {3:0.####} l\n",
                x1, Height - y1, x2, Height - y2);
        }

        public void DrawLines(Color color, double lineWidth, (double X, double Y)[] points)
        {
            if (points == null || points.Length < 2 || color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth);
            _currentCommands.AppendFormat(CultureInfo.InvariantCulture,
                "{0:0.####} {1:0.####} m\n",
                points[0].X, Height - points[0].Y);

            for (int i = 1; i < points.Length; i++)
            {
                _currentCommands.AppendFormat(CultureInfo.InvariantCulture,
                    "{0:0.####} {1:0.####} l\n",
                    points[i].X, Height - points[i].Y);
            }
        }

        public void DrawString(string text, MinimalPdfFont font, double fontSize, Color color, double x, double y)
        {
            if (string.IsNullOrEmpty(text) || color.A == 0)
                return;

            FlushStrokePath();

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
            FlushStrokePath();

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
