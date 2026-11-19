using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace SharpLayout.NativePdf
{
    public readonly struct LayoutSize
    {
        public double Width { get; }
        public double Height { get; }

        public LayoutSize(double width, double height)
        {
            Width = width;
            Height = height;
        }
    }

    public class TrueTypeMetrics
    {
        public byte[] TtfBytes { get; }
        private readonly bool _simulateBold;
        private const int FastLookupSize = 0x0500;

        /// <summary>
        /// Factor of the em size of a regular font to look bold. Used for bold simulation.
        /// Value of 2% found in original XPS 1.0 documentation.
        /// Source: https://github.com/AVPolyakov/PDFsharp/blob/master/src/PdfSharp/Drawing/FontHelper.cs#L104-L109
        /// </summary>
        private const double BoldEmphasis = 0.02;
        
        public string FontFamily { get; private set; }
        public bool IsBold { get; private set; }
        public bool IsItalic { get; private set; }

        public ushort UnitsPerEm { get; private set; }
        public short Ascender { get; private set; }
        public short Descender { get; private set; }
        public short LineGap { get; private set; }
        public ushort WinAscent { get; private set; }
        public ushort WinDescent { get; private set; }
        public short CapHeight { get; private set; }
        public short XMin { get; private set; }
        public short YMin { get; private set; }
        public short XMax { get; private set; }
        public short YMax { get; private set; }
        public bool IsSymbol { get; private set; }

        private readonly ushort[] _fastGlyphIds = new ushort[FastLookupSize];
        private readonly Dictionary<int, ushort> _fallbackGlyphIds = new Dictionary<int, ushort>();
        private ushort[] _advanceWidths;

        public TrueTypeMetrics(byte[] ttfBytes, bool simulateBold = false)
        {
            TtfBytes = ttfBytes;
            _simulateBold = simulateBold;
            ParseFont(ttfBytes);
        }

        public LayoutSize MeasureString(string text, double fontSize)
        {
            double lineHeight = WinAscent > 0
                ? (double)(WinAscent + WinDescent) / UnitsPerEm * fontSize
                : (double)(Ascender - Descender + LineGap) / UnitsPerEm * fontSize;

            if (string.IsNullOrEmpty(text))
                return new LayoutSize(0, lineHeight);

            long totalAdvanceUnits = 0;
            ushort[] fastGlyphs = _fastGlyphIds;
            ushort[] widths = _advanceWidths;
            int widthsLen = widths.Length;
            ushort fallbackWidth = widthsLen > 0 ? widths[widthsLen - 1] : (ushort)0;

            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch < 32)
                    continue;

                ushort gid;
                if (ch < FastLookupSize)
                {
                    gid = fastGlyphs[ch];
                }
                else if (char.IsHighSurrogate(ch) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    int codePoint = char.ConvertToUtf32(ch, text[i + 1]);
                    i++;
                    gid = _fallbackGlyphIds.TryGetValue(codePoint, out ushort found) ? found : (ushort)0;
                }
                else
                {
                    gid = _fallbackGlyphIds.TryGetValue(ch, out ushort found) ? found : (ushort)0;
                }

                totalAdvanceUnits += gid < widthsLen ? widths[gid] : fallbackWidth;
            }

            double width = (double)totalAdvanceUnits / UnitsPerEm * fontSize;

            // Exact replica of XStyleSimulations.BoldSimulation from PDFsharp (+2% of font size per character)
            if (_simulateBold)
            {
                width += text.Length * fontSize * BoldEmphasis;
            }

            return new LayoutSize(width, lineHeight);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public ushort GetGlyphId(int codePoint)
        {
            if ((uint)codePoint < FastLookupSize)
                return _fastGlyphIds[codePoint];

            return _fallbackGlyphIds.TryGetValue(codePoint, out ushort gid) ? gid : (ushort)0;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public ushort GetAdvanceWidth(ushort glyphId)
        {
            ushort[] widths = _advanceWidths;
            if (widths == null || widths.Length == 0)
                return 0;

            return glyphId < widths.Length ? widths[glyphId] : widths[widths.Length - 1];
        }

        private void SetGlyphMapping(int codePoint, ushort glyphId)
        {
            if (glyphId == 0) return;

            if ((uint)codePoint < FastLookupSize)
                _fastGlyphIds[codePoint] = glyphId;
            else
                _fallbackGlyphIds[codePoint] = glyphId;
        }

        private void ParseFont(byte[] ttfBytes)
        {
            using (var ms = new MemoryStream(ttfBytes))
            {
                ReadUInt32(ms); // sfntVersion
                ushort numTables = ReadUInt16(ms);
                ms.Position += 6;

                var tables = new Dictionary<string, (uint Offset, uint Length)>();
                for (int i = 0; i < numTables; i++)
                {
                    byte[] tagBytes = ReadBytes(ms, 4);
                    string tag = Encoding.ASCII.GetString(tagBytes);

                    ReadUInt32(ms); // checksum
                    uint offset = ReadUInt32(ms);
                    uint length = ReadUInt32(ms);
                    tables[tag] = (offset, length);
                }

                // 1. Таблица head
                ms.Position = tables["head"].Offset + 18;
                UnitsPerEm = ReadUInt16(ms);

                ms.Position = tables["head"].Offset + 36;
                XMin = ReadInt16(ms);
                YMin = ReadInt16(ms);
                XMax = ReadInt16(ms);
                YMax = ReadInt16(ms);
                
                // 2. Таблица hhea
                ms.Position = tables["hhea"].Offset + 4;
                Ascender = ReadInt16(ms);
                Descender = ReadInt16(ms);
                LineGap = ReadInt16(ms);
                ms.Position += 24;
                ushort numberOfHMetrics = ReadUInt16(ms);

                // 3. Таблица OS/2 (флаги стиля и метрики высоты)
                if (tables.TryGetValue("OS/2", out var os2Table))
                {
                    ms.Position = os2Table.Offset;
                    ushort os2Version = ReadUInt16(ms);

                    ms.Position = os2Table.Offset + 62;
                    ushort fsSelection = ReadUInt16(ms);
                    IsItalic = (fsSelection & 0x0001) != 0;
                    IsBold = (fsSelection & 0x0020) != 0;

                    ms.Position = os2Table.Offset + 74;
                    WinAscent = ReadUInt16(ms);
                    WinDescent = ReadUInt16(ms);

                    if (os2Version >= 2 && os2Table.Length >= 90)
                    {
                        ms.Position = os2Table.Offset + 88;
                        CapHeight = ReadInt16(ms);
                    }
                }

                // 4. Таблица name (FontFamily)
                if (tables.TryGetValue("name", out var nameTable))
                {
                    FontFamily = ParseFontFamily(ms, nameTable.Offset);
                }

                // 5. Таблица maxp
                ms.Position = tables["maxp"].Offset + 4;
                ushort numGlyphs = ReadUInt16(ms);

                // 6. Таблица hmtx
                ms.Position = tables["hmtx"].Offset;
                _advanceWidths = new ushort[numGlyphs];
                ushort lastAdvance = 0;
                for (int i = 0; i < numberOfHMetrics; i++)
                {
                    lastAdvance = ReadUInt16(ms);
                    ReadInt16(ms); // leftSideBearing
                    _advanceWidths[i] = lastAdvance;
                }
                for (int i = numberOfHMetrics; i < numGlyphs; i++)
                {
                    _advanceWidths[i] = lastAdvance;
                }

                // 7. Таблица cmap
                ParseCmapFormat4(ms, tables["cmap"].Offset);
            }
        }

        private static string ParseFontFamily(MemoryStream ms, uint nameOffset)
        {
            ms.Position = nameOffset + 2;
            ushort count = ReadUInt16(ms);
            ushort stringOffset = ReadUInt16(ms);
            long storageBase = nameOffset + stringOffset;

            string fallback = null;
            for (int i = 0; i < count; i++)
            {
                ushort platformId = ReadUInt16(ms);
                ushort encodingId = ReadUInt16(ms);
                ushort languageId = ReadUInt16(ms);
                ushort nameId = ReadUInt16(ms);
                ushort length = ReadUInt16(ms);
                ushort offset = ReadUInt16(ms);

                if (nameId == 1) // Font Family name
                {
                    long savedPos = ms.Position;
                    ms.Position = storageBase + offset;
                    byte[] raw = ReadBytes(ms, length);
                    ms.Position = savedPos;

                    string value = null;
                    if (platformId == 3 || platformId == 0)
                        value = Encoding.BigEndianUnicode.GetString(raw);
                    else if (platformId == 1 && encodingId == 0)
                        value = Encoding.ASCII.GetString(raw);

                    if (value != null)
                    {
                        if (platformId == 3 && languageId == 0x0409)
                            return value;

                        if (fallback == null)
                            fallback = value;
                    }
                }
            }
            return fallback;
        }

        private void ParseCmapFormat4(MemoryStream ms, uint cmapOffset)
        {
            ms.Position = cmapOffset + 2;
            ushort numSubtables = ReadUInt16(ms);
            uint targetSubtableOffset = 0;
            bool isSymbolEncoding = false;

            for (int i = 0; i < numSubtables; i++)
            {
                ushort platformId = ReadUInt16(ms);
                ushort encodingId = ReadUInt16(ms);
                uint subtableOffset = ReadUInt32(ms);

                if (platformId == 3 && encodingId == 1) // Windows Unicode BMP
                {
                    targetSubtableOffset = cmapOffset + subtableOffset;
                    isSymbolEncoding = false;
                    break;
                }
                if (platformId == 3 && encodingId == 0) // Windows Symbol (Wingdings, Symbol, Webdings)
                {
                    targetSubtableOffset = cmapOffset + subtableOffset;
                    isSymbolEncoding = true;
                }
            }
            
            IsSymbol = isSymbolEncoding;

            if (targetSubtableOffset == 0)
                throw new InvalidDataException("В шрифте не найдена таблица cmap (Windows Unicode BMP / Symbol).");

            ms.Position = targetSubtableOffset;
            ushort format = ReadUInt16(ms);
            if (format != 4)
                throw new NotSupportedException($"Ожидался формат cmap 4, но получен формат {format}.");

            ms.Position += 4;
            ushort segCount = (ushort)(ReadUInt16(ms) / 2);
            ms.Position += 6;

            ushort[] endCode = new ushort[segCount];
            for (int i = 0; i < segCount; i++) endCode[i] = ReadUInt16(ms);

            ms.Position += 2;

            ushort[] startCode = new ushort[segCount];
            for (int i = 0; i < segCount; i++) startCode[i] = ReadUInt16(ms);

            short[] idDelta = new short[segCount];
            for (int i = 0; i < segCount; i++) idDelta[i] = ReadInt16(ms);

            long idRangeOffsetTablePos = ms.Position;
            ushort[] idRangeOffset = new ushort[segCount];
            for (int i = 0; i < segCount; i++) idRangeOffset[i] = ReadUInt16(ms);

            for (int s = 0; s < segCount; s++)
            {
                if (startCode[s] == 0xFFFF)
                    break;
                for (int c = startCode[s]; c <= endCode[s]; c++)
                {
                    ushort gid;
                    if (idRangeOffset[s] == 0)
                    {
                        gid = (ushort)((c + idDelta[s]) & 0xFFFF);
                    }
                    else
                    {
                        long glyphIndexPos = idRangeOffsetTablePos + (s * 2) + idRangeOffset[s] + ((c - startCode[s]) * 2);
                        ms.Position = glyphIndexPos;
                        gid = ReadUInt16(ms);
                        if (gid != 0)
                            gid = (ushort)((gid + idDelta[s]) & 0xFFFF);
                    }

                    SetGlyphMapping(c, gid);

                    // Для символьных шрифтов (Wingdings 2) дублируем коды из 0xF000..0xF0FF в 0x0000..0x00FF
                    if (isSymbolEncoding && c >= 0xF000 && c <= 0xF0FF)
                    {
                        SetGlyphMapping(c - 0xF000, gid);
                    }
                }
            }
        }
        
        private static byte[] ReadBytes(Stream s, int count)
        {
            byte[] buffer = new byte[count];
            int read = 0;
            while (read < count)
            {
                int n = s.Read(buffer, read, count - read);
                if (n <= 0) throw new EndOfStreamException();
                read += n;
            }
            return buffer;
        }
        private static ushort ReadUInt16(Stream s)
        {
            int b0 = s.ReadByte();
            int b1 = s.ReadByte();
            return (ushort)((b0 << 8) | b1);
        }
        private static short ReadInt16(Stream s) => (short)ReadUInt16(s);
        private static uint ReadUInt32(Stream s)
        {
            uint b0 = (uint)s.ReadByte();
            uint b1 = (uint)s.ReadByte();
            uint b2 = (uint)s.ReadByte();
            uint b3 = (uint)s.ReadByte();
            return (b0 << 24) | (b1 << 16) | (b2 << 8) | b3;
        }
    }
}