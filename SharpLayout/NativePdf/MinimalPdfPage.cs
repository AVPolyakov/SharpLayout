using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using PdfSharp;

namespace SharpLayout.NativePdf
{
    public class MinimalPdfPage
    {
        private const string HexChars = "0123456789ABCDEF";

        public double Width => Orientation switch
        {
            PageOrientation.Portrait => 595,
            PageOrientation.Landscape => 842,
            _ => throw new ArgumentOutOfRangeException()
        };

        public double Height => Orientation switch
        {
            PageOrientation.Portrait => 842,
            PageOrientation.Landscape => 595,
            _ => throw new ArgumentOutOfRangeException()
        };
        
        public PageOrientation Orientation { get; set; }

        internal Dictionary<byte, string> AlphaStates { get; } = new Dictionary<byte, string>();

        private readonly List<(string Prefix, MinimalPdfFont Font, string Text)> _chunks =
            new List<(string, MinimalPdfFont, string)>();
        private readonly StringBuilder _cmd = new StringBuilder();

        // Состояние векторного пути (линии)
        private bool _hasOpenPath;
        private double _currentStrokeWidth = -1;
        private Color _currentStrokeColor;
        private string _currentDashPattern = "";

        // Состояние текстового блока, заливки и прозрачности
        private bool _inTextObject;
        private bool _isBoldSimulated;
        private bool _hasFillColor;
        private Color _currentFillColor;
        private byte _currentAlpha = 255;
        private string _currentFontAlias;
        private double _currentFontSize = -1;
        private double _textCursorX;
        private double _textCursorY;
        
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
                if (_isBoldSimulated)
                {
                    _cmd.Append("0 Tr 0 Tc\n");
                    _isBoldSimulated = false;
                }
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

        private void EnsureStrokeStyle(Color color, double lineWidth, double[] dashPattern = null)
        {
            EndTextObject();
            
            string patternKey = FormatDashPattern(dashPattern, lineWidth);

            bool styleChanged = !_hasOpenPath
                || Math.Abs(_currentStrokeWidth - lineWidth) > 0.0001
                || _currentStrokeColor != color
                || _currentDashPattern != patternKey;

            if (styleChanged)
            {
                FlushStrokePath();
                _currentStrokeWidth = lineWidth;
                _currentStrokeColor = color;

                if (_currentDashPattern != patternKey)
                {
                    _currentDashPattern = patternKey;
                    _cmd.Append('[');
                    _cmd.Append(patternKey);
                    _cmd.Append("] 0 d ");
                }
                
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

        private static string FormatDashPattern(double[] dashPattern, double lineWidth)
        {
            if (dashPattern == null || dashPattern.Length == 0)
                return "";

            var sb = new StringBuilder();
            for (int i = 0; i < dashPattern.Length; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append((dashPattern[i] * lineWidth).ToString("0.####", CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }

        public void DrawLine(Color color, double lineWidth, double x1, double y1, double x2, double y2, double[] dashPattern = null)
         {
            if (color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth, dashPattern);
            AppendNum(x1);
            _cmd.Append(' ');
            AppendNum(Height - y1);
            _cmd.Append(" m ");
            AppendNum(x2);
            _cmd.Append(' ');
            AppendNum(Height - y2);
            _cmd.Append(" l\n");
        }

        public void DrawLines(Color color, double lineWidth, (double X, double Y)[] points, double[] dashPattern = null)
        {
            if (points == null || points.Length < 2 || color.A == 0 || lineWidth <= 0)
                return;

            EnsureStrokeStyle(color, lineWidth, dashPattern);
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

            bool simulateBold = font.Metrics.SimulateBold;
            if (simulateBold)
            {
                double emphasis = fontSize * TrueTypeMetrics.BoldEmphasis;
                AppendNum(emphasis);
                _cmd.Append(" w ");
                AppendColorComponent(color.R);
                _cmd.Append(' ');
                AppendColorComponent(color.G);
                _cmd.Append(' ');
                AppendColorComponent(color.B);
                _cmd.Append(" RG 2 Tr ");
                AppendNum(emphasis);
                _cmd.Append(" Tc\n");
                _isBoldSimulated = true;
                _currentStrokeWidth = -1;
            }
            else if (_isBoldSimulated)
            {
                _cmd.Append("0 Tr 0 Tc\n");
                _isBoldSimulated = false;
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
        {
            // Быстрый без-аллокационный путь для стандартных координат PDF с точностью до 4 знаков (0.####)
            if (val >= -200000.0 && val <= 200000.0)
            {
                long scaled = (long)Math.Round(val * 10000.0, MidpointRounding.AwayFromZero);
                if (scaled == 0)
                {
                    _cmd.Append('0');
                    return;
                }

                if (scaled < 0)
                {
                    _cmd.Append('-');
                    scaled = -scaled;
                }

                long intPart = scaled / 10000;
                int fracPart = (int)(scaled % 10000);

                _cmd.Append(intPart);

                if (fracPart > 0)
                {
                    _cmd.Append('.');
                    int d1 = fracPart / 1000;
                    int d2 = (fracPart / 100) % 10;
                    int d3 = (fracPart / 10) % 10;
                    int d4 = fracPart % 10;

                    _cmd.Append((char)('0' + d1));
                    if (d2 != 0 || d3 != 0 || d4 != 0)
                    {
                        _cmd.Append((char)('0' + d2));
                        if (d3 != 0 || d4 != 0)
                        {
                            _cmd.Append((char)('0' + d3));
                            if (d4 != 0)
                            {
                                _cmd.Append((char)('0' + d4));
                            }
                        }
                    }
                }
                return;
            }

            _cmd.Append(val.ToString("0.####", CultureInfo.InvariantCulture));
        }

        private void AppendColorComponent(byte c)
        {
            if (c == 0)
            {
                _cmd.Append('0');
                return;
            }
            if (c == 255)
            {
                _cmd.Append('1');
                return;
            }

            // Точность 0.### (как (c / 255.0).ToString("0.###")) без создания строк
            int scaled = (int)Math.Round(c * 1000.0 / 255.0, MidpointRounding.AwayFromZero);
            if (scaled >= 1000)
            {
                _cmd.Append('1');
                return;
            }

            _cmd.Append("0.");
            int d1 = scaled / 100;
            int d2 = (scaled / 10) % 10;
            int d3 = scaled % 10;

            _cmd.Append((char)('0' + d1));
            if (d2 != 0 || d3 != 0)
            {
                _cmd.Append((char)('0' + d2));
                if (d3 != 0)
                {
                    _cmd.Append((char)('0' + d3));
                }
            }
        }
        
        internal HashSet<MinimalPdfImage> UsedImages { get; } = new HashSet<MinimalPdfImage>();

        public void DrawImage(MinimalPdfImage image, double x, double y, double width, double height)
        {
            if (image == null || width <= 0 || height <= 0)
                return;

            EndTextObject();
            FlushStrokePath();
            UsedImages.Add(image);

            double bottomY = Height - (y + height);

            _cmd.Append("q ");
            AppendNum(width);
            _cmd.Append(" 0 0 ");
            AppendNum(height);
            _cmd.Append(' ');
            AppendNum(x);
            _cmd.Append(' ');
            AppendNum(bottomY);
            _cmd.Append(" cm /");
            _cmd.Append(image.Alias);
            _cmd.Append(" Do Q\n");
        }        
    }
}
