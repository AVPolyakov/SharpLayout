using System;
using System.Collections.Generic;
using System.IO;
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
        public ushort UnitsPerEm { get; private set; }
        public short Ascender { get; private set; }
        public short Descender { get; private set; }
        public short LineGap { get; private set; }
        public ushort WinAscent { get; private set; }
        public ushort WinDescent { get; private set; }

        private readonly Dictionary<int, ushort> _charToGlyph = new Dictionary<int, ushort>();
        private ushort[] _advanceWidths;

        public TrueTypeMetrics(byte[] ttfBytes)
        {
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
            for (int i = 0; i < text.Length; i++)
            {
                int codePoint = char.ConvertToUtf32(text, i);
                if (char.IsSurrogatePair(text, i))
                    i++;

                ushort glyphId = GetGlyphId(codePoint);
                totalAdvanceUnits += GetAdvanceWidth(glyphId);
            }

            double width = (double)totalAdvanceUnits / UnitsPerEm * fontSize;
            return new LayoutSize(width, lineHeight);
        }

        public ushort GetGlyphId(int codePoint)
        {
            return _charToGlyph.TryGetValue(codePoint, out ushort gid) ? gid : (ushort)0;
        }

        public ushort GetAdvanceWidth(ushort glyphId)
        {
            if (_advanceWidths == null || _advanceWidths.Length == 0)
                return 0;

            return glyphId < _advanceWidths.Length
                ? _advanceWidths[glyphId]
                : _advanceWidths[_advanceWidths.Length - 1];
        }

        private void ParseFont(byte[] ttfBytes)
        {
            using (var ms = new MemoryStream(ttfBytes))
            {
                ReadUInt32(ms); // sfntVersion
                ushort numTables = ReadUInt16(ms);
                ms.Position += 6; // searchRange, entrySelector, rangeShift

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

                // 1. Таблица head: получаем UnitsPerEm (смещение +18 байт от начала таблицы)
                ms.Position = tables["head"].Offset + 18;
                UnitsPerEm = ReadUInt16(ms);

                // 2. Таблица hhea: получаем Ascender, Descender, LineGap и число метрик
                ms.Position = tables["hhea"].Offset + 4;
                Ascender = ReadInt16(ms);
                Descender = ReadInt16(ms);
                LineGap = ReadInt16(ms);
                ms.Position += 24; // пропускаем промежуточные поля заголовка hhea
                ushort numberOfHMetrics = ReadUInt16(ms);

                // 3. Таблица OS/2: получаем usWinAscent и usWinDescent для расчёта высоты строки как в GDI+/PDFsharp
                if (tables.TryGetValue("OS/2", out var os2Table))
                {
                    ms.Position = os2Table.Offset + 74;
                    WinAscent = ReadUInt16(ms);
                    WinDescent = ReadUInt16(ms);
                }

                // 4. Таблица maxp: узнаём общее количество глифов в шрифте
                ms.Position = tables["maxp"].Offset + 4;
                ushort numGlyphs = ReadUInt16(ms);

                // 5. Таблица hmtx: считываем ширину каждого глифа
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

                // 6. Таблица cmap: строим карту Unicode -> GlyphID
                ParseCmapFormat4(ms, tables["cmap"].Offset);
            }
        }

        private void ParseCmapFormat4(MemoryStream ms, uint cmapOffset)
        {
            ms.Position = cmapOffset + 2;
            ushort numSubtables = ReadUInt16(ms);
            uint targetSubtableOffset = 0;

            for (int i = 0; i < numSubtables; i++)
            {
                ushort platformId = ReadUInt16(ms);
                ushort encodingId = ReadUInt16(ms);
                uint subtableOffset = ReadUInt32(ms);

                // Стандартная таблица Windows Unicode BMP (Platform 3, Encoding 1)
                if (platformId == 3 && encodingId == 1)
                {
                    targetSubtableOffset = cmapOffset + subtableOffset;
                    break;
                }
            }

            if (targetSubtableOffset == 0)
                throw new InvalidDataException("В шрифте не найдена таблица cmap (Windows Unicode BMP).");

            ms.Position = targetSubtableOffset;
            ushort format = ReadUInt16(ms);
            if (format != 4)
                throw new NotSupportedException($"Ожидался формат cmap 4, но получен формат {format}.");

            ms.Position += 4; // length, language
            ushort segCount = (ushort)(ReadUInt16(ms) / 2);
            ms.Position += 6; // searchRange, entrySelector, rangeShift

            ushort[] endCode = new ushort[segCount];
            for (int i = 0; i < segCount; i++) endCode[i] = ReadUInt16(ms);

            ms.Position += 2; // reservedPad

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

                    if (gid != 0)
                        _charToGlyph[c] = gid;
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
