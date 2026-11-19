using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SharpLayout.NativePdf
{
    public class MinimalPdfPage
    {
        private const string HexChars = "0123456789ABCDEF";

        public double Width { get; }
        public double Height { get; }

        internal Dictionary<byte, string> AlphaStates { get; } = new Dictionary<byte, string>();

        private readonly List<(string Prefix, MinimalPdfFont Font, string Text)> _chunks =
            new List<(string, MinimalPdfFont, string)>();
        private readonly StringBuilder _cmd = new StringBuilder();

        // Состояние векторного пути (линии)
        private bool _hasOpenPath;
        private double _currentStrokeWidth = -1;
        private Color _currentStrokeColor;

        // Состояние текстового блока, заливки и прозрачности
        private bool _inTextObject;
        private bool _hasFillColor;
        private Color _currentFillColor;
        private byte _currentAlpha = 255;
        private string _currentFontAlias;
        private double _currentFontSize = -1;
        private double _textCursorX;
        private double _textCursorY;

        public MinimalPdfPage(double width, double height)
        {
            Width = width;
            Height = height;
        }

        private string GetOrRegisterAlphaState(byte alpha)
        {
            if (!AlphaStates.TryGetValue(alpha, out string alias))
            {
                alias = $"GSA{alpha}";
                AlphaStates[alpha] = alias;
            }
            return alias;
        }

        private void EnsureAlpha(byte alpha)
        {
            if (_currentAlpha != alpha)
            {
                _currentAlpha = alpha;
                string gsAlias = GetOrRegisterAlphaState(alpha);
                _cmd.Append('/');
                _cmd.Append(gsAlias);
                _cmd.Append(" gs\n");
            }
        }

        private void FlushStrokePath()
        {
            if (_hasOpenPath)
            {
                _cmd.Append("S\n");
                _hasOpenPath = false;
            }
        }

        private void EndTextObject()
        {
            if (_inTextObject)
            {
                _cmd.Append("ET\n");
                _inTextObject = false;
            }
        }

        private void EnsureFillColor(Color color)
        {
            EnsureAlpha(color.A);

            if (!_hasFillColor || _currentFillColor.R != color.R || _currentFillColor.G != color.G || _currentFillColor.B != color.B)
            {
                _currentFillColor = color;
                _hasFillColor = true;
                AppendColorComponent(color.R);
                _cmd.Append(' ');
                AppendColorComponent(color.G);
                _cmd.Append(' ');
                AppendColorComponent(color.B);
                _cmd.Append(" rg\n");
            }
        }

        private void EnsureStrokeStyle(Color color, double lineWidth)
        {
            EndTextObject();

            bool styleChanged = !_hasOpenPath
                || Math.Abs(_currentStrokeWidth - lineWidth) > 0.0001
                || _currentStrokeColor != color;

            if (styleChanged)
            {
                FlushStrokePath();
                _currentStrokeWidth = lineWidth;
                _currentStrokeColor = color;

                EnsureAlpha(color.A);

                AppendNum(lineWidth);
                _cmd.Append(" w ");
                AppendColorComponent(color.R);
                _cmd.Append(' ');
                AppendColorComponent(color.G);
                _cmd.Append(' ');
                AppendColorComponent(color.B);
                _cmd.Append(" RG\n");

                _hasOpenPath = true;
            }
        }

        public void DrawLine(Color color, double lineWidth, double x1, double y1, double x2, double y2)
        {
            if (color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth);
            AppendNum(x1);
            _cmd.Append(' ');
            AppendNum(Height - y1);
            _cmd.Append(" m ");
            AppendNum(x2);
            _cmd.Append(' ');
            AppendNum(Height - y2);
            _cmd.Append(" l\n");
        }

        public void DrawLines(Color color, double lineWidth, (double X, double Y)[] points)
        {
            if (points == null || points.Length < 2 || color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth);
            AppendNum(points[0].X);
            _cmd.Append(' ');
            AppendNum(Height - points[0].Y);
            _cmd.Append(" m\n");

            for (int i = 1; i < points.Length; i++)
            {
                AppendNum(points[i].X);
                _cmd.Append(' ');
                AppendNum(Height - points[i].Y);
                _cmd.Append(" l\n");
            }
        }

        public void DrawRectangle(Color fillColor, double x, double y, double width, double height)
        {
            if (fillColor.A == 0 || width <= 0 || height <= 0)
                return;

            EndTextObject();
            FlushStrokePath();
            EnsureFillColor(fillColor);

            double bottomY = Height - (y + height);

            AppendNum(x);
            _cmd.Append(' ');
            AppendNum(bottomY);
            _cmd.Append(' ');
            AppendNum(width);
            _cmd.Append(' ');
            AppendNum(height);
            _cmd.Append(" re f\n");
        }

        public void DrawString(string text, MinimalPdfFont font, double fontSize, Color color, double x, double y)
        {
            if (string.IsNullOrEmpty(text) || color.A == 0)
                return;

            FlushStrokePath();

            double baselineY = Height - y;

            if (!_inTextObject)
            {
                _cmd.Append("BT\n");
                _inTextObject = true;
                _textCursorX = 0;
                _textCursorY = 0;
            }

            EnsureFillColor(color);

            if (_currentFontAlias != font.Alias || Math.Abs(_currentFontSize - fontSize) > 0.0001)
            {
                _currentFontAlias = font.Alias;
                _currentFontSize = fontSize;
                _cmd.Append('/');
                _cmd.Append(font.Alias);
                _cmd.Append(' ');
                AppendNum(fontSize);
                _cmd.Append(" Tf\n");
            }

            double dx = x - _textCursorX;
            double dy = baselineY - _textCursorY;
            _textCursorX = x;
            _textCursorY = baselineY;

            AppendNum(dx);
            _cmd.Append(' ');
            AppendNum(dy);
            _cmd.Append(" Td <");

            for (int i = 0; i < text.Length; i++)
            {
                int cp = char.ConvertToUtf32(text, i);
                if (char.IsSurrogatePair(text, i)) i++;
                font.UsedCodePoints.Add(cp);
            }

            string prefix = _cmd.ToString();
            _cmd.Clear();

            _chunks.Add((prefix, font, text));
            _cmd.Append("> Tj\n");
        }

        internal string GetContent()
        {
            EndTextObject();
            FlushStrokePath();

            var sb = new StringBuilder(_chunks.Count * 48 + _cmd.Length);
            for (int c = 0; c < _chunks.Count; c++)
            {
                var chunk = _chunks[c];
                sb.Append(chunk.Prefix);

                string text = chunk.Text;
                MinimalPdfFont font = chunk.Font;
                for (int i = 0; i < text.Length; i++)
                {
                    int cp = char.ConvertToUtf32(text, i);
                    if (char.IsSurrogatePair(text, i)) i++;

                    ushort gid = font.GetFinalGlyphId(cp);
                    sb.Append(HexChars[(gid >> 12) & 0xF]);
                    sb.Append(HexChars[(gid >> 8) & 0xF]);
                    sb.Append(HexChars[(gid >> 4) & 0xF]);
                    sb.Append(HexChars[gid & 0xF]);
                }
            }

            if (_cmd.Length > 0)
                sb.Append(_cmd);

            return sb.ToString();
        }

        private void AppendNum(double val)
            => _cmd.Append(val.ToString("0.####", CultureInfo.InvariantCulture));

        private void AppendColorComponent(byte c)
            => _cmd.Append((c / 255.0).ToString("0.###", CultureInfo.InvariantCulture));
    }
}
