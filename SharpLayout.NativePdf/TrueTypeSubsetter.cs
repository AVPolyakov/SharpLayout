using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharpLayout.NativePdf
{
    public static class TrueTypeSubsetter
    {
        public static byte[] CreateSubset(byte[] ttfBytes, IEnumerable<ushort> usedGlyphIds)
        {
            var tables = new Dictionary<string, byte[]>();
            short indexToLocFormat;
            ushort numGlyphs;

            // 1. Разбираем исходный TTF на отдельные таблицы
            using (var ms = new MemoryStream(ttfBytes))
            {
                ReadUInt32(ms); // sfntVersion
                ushort numTables = ReadUInt16(ms);
                ms.Position += 6;

                var entries = new List<(string Tag, uint Offset, uint Length)>();
                for (int i = 0; numTables > i; i++)
                {
                    string tag = Encoding.ASCII.GetString(ReadBytes(ms, 4));
                    ReadUInt32(ms); // checksum
                    uint offset = ReadUInt32(ms);
                    uint length = ReadUInt32(ms);
                    entries.Add((tag, offset, length));
                }

                foreach (var entry in entries)
                {
                    ms.Position = entry.Offset;
                    tables[entry.Tag] = ReadBytes(ms, (int)entry.Length);
                }
            }

            // Читаем формат таблицы loca (offset +50 в таблице head) и число глифов (offset +4 в maxp)
            indexToLocFormat = ReadInt16BE(tables["head"], 50);
            numGlyphs = ReadUInt16BE(tables["maxp"], 4);

            // 2. Читаем старые смещения глифов из loca
            uint[] oldOffsets = new uint[numGlyphs + 1];
            byte[] locaData = tables["loca"];
            for (int i = 0; numGlyphs >= i; i++)
            {
                oldOffsets[i] = indexToLocFormat == 0
                    ? (uint)(ReadUInt16BE(locaData, i * 2) * 2)
                    : ReadUInt32BE(locaData, i * 4);
            }

            // 3. Собираем все нужные глифы (включая компоненты составных глифов вроде 'ё' и 'й')
            var keepGlyphs = new HashSet<ushort>(usedGlyphIds) { 0 }; // 0 (.notdef) обязателен
            byte[] glyfData = tables["glyf"];
            var queue = new Queue<ushort>(keepGlyphs);

            while (queue.Count > 0)
            {
                ushort gid = queue.Dequeue();
                if (gid >= numGlyphs) continue;

                uint start = oldOffsets[gid];
                uint end = oldOffsets[gid + 1];
                if (end <= start || start + 10 > glyfData.Length) continue;

                short numberOfContours = ReadInt16BE(glyfData, (int)start);
                if (0 > numberOfContours) // Составной глиф (Composite)
                {
                    int pos = (int)start + 10;
                    const ushort ARG_1_AND_2_ARE_WORDS = 0x0001;
                    const ushort WE_HAVE_A_SCALE = 0x0008;
                    const ushort WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
                    const ushort WE_HAVE_A_TWO_BY_TWO = 0x0080;
                    const ushort MORE_COMPONENTS = 0x0020;

                    while (pos + 4 <= end)
                    {
                        ushort flags = ReadUInt16BE(glyfData, pos);
                        ushort componentGid = ReadUInt16BE(glyfData, pos + 2);
                        pos += 4;

                        if (keepGlyphs.Add(componentGid))
                            queue.Enqueue(componentGid);

                        pos += (flags & ARG_1_AND_2_ARE_WORDS) != 0 ? 4 : 2;
                        if ((flags & WE_HAVE_A_SCALE) != 0) pos += 2;
                        else if ((flags & WE_HAVE_AN_X_AND_Y_SCALE) != 0) pos += 4;
                        else if ((flags & WE_HAVE_A_TWO_BY_TWO) != 0) pos += 8;

                        if ((flags & MORE_COMPONENTS) == 0) break;
                    }
                }
            }

            // 4. Пересобираем таблицы glyf и loca, оставляя контуры только для keepGlyphs
            uint[] newOffsets = new uint[numGlyphs + 1];
            using (var newGlyf = new MemoryStream())
            {
                for (ushort gid = 0; numGlyphs > gid; gid++)
                {
                    newOffsets[gid] = (uint)newGlyf.Position;
                    if (keepGlyphs.Contains(gid))
                    {
                        uint start = oldOffsets[gid];
                        uint length = oldOffsets[gid + 1] - start;
                        if (length > 0 && start + length <= glyfData.Length)
                        {
                            newGlyf.Write(glyfData, (int)start, (int)length);
                            // Выравнивание по 4 байтам
                            while ((newGlyf.Position & 3) != 0) newGlyf.WriteByte(0);
                        }
                    }
                }
                newOffsets[numGlyphs] = (uint)newGlyf.Position;
                tables["glyf"] = newGlyf.ToArray();
            }

            // Записываем новую таблицу loca (используем длинный формат 32-bit, обновив head)
            byte[] newLoca = new byte[(numGlyphs + 1) * 4];
            for (int i = 0; numGlyphs >= i; i++)
            {
                WriteUInt32BE(newLoca, i * 4, newOffsets[i]);
            }
            tables["loca"] = newLoca;

            // Указываем в head, что формат loca теперь 1 (32-bit)
            byte[] newHead = (byte[])tables["head"].Clone();
            newHead[50] = 0;
            newHead[51] = 1;
            tables["head"] = newHead;

            // Удаляем лишние таблицы, не нужные для рендеринга в PDF (hinting, GSUB, GPOS, DSIG и т.д.)
            string[] requiredTags = { "head", "hhea", "maxp", "OS/2", "hmtx", "cmap", "loca", "glyf", "cvt ", "fpgm", "prep" };
            var finalTables = tables
                .Where(kv => requiredTags.Contains(kv.Key))
                .OrderBy(kv => kv.Key, StringComparer.Ordinal)
                .ToList();

            // 5. Собираем итоговый бинарный TTF-файл
            return BuildTtfBinary(finalTables);
        }

        private static byte[] BuildTtfBinary(List<KeyValuePair<string, byte[]>> tables)
        {
            ushort numTables = (ushort)tables.Count;
            ushort entrySelector = 0;
            ushort searchRange = 1;
            while (numTables >= searchRange * 2)
            {
                searchRange *= 2;
                entrySelector++;
            }
            searchRange *= 16;
            ushort rangeShift = (ushort)(numTables * 16 - searchRange);

            using (var ms = new MemoryStream())
            {
                WriteUInt32(ms, 0x00010000); // TrueType version 1.0
                WriteUInt16(ms, numTables);
                WriteUInt16(ms, searchRange);
                WriteUInt16(ms, entrySelector);
                WriteUInt16(ms, rangeShift);

                uint currentOffset = (uint)(12 + numTables * 16);
                foreach (var kv in tables)
                {
                    byte[] tagBytes = Encoding.ASCII.GetBytes(kv.Key.PadRight(4));
                    ms.Write(tagBytes, 0, 4);
                    WriteUInt32(ms, 0); // checksum (PDF-ридеры не требуют)
                    WriteUInt32(ms, currentOffset);
                    WriteUInt32(ms, (uint)kv.Value.Length);

                    uint paddedLen = (uint)((kv.Value.Length + 3) & ~3);
                    currentOffset += paddedLen;
                }

                foreach (var kv in tables)
                {
                    ms.Write(kv.Value, 0, kv.Value.Length);
                    while ((ms.Position & 3) != 0) ms.WriteByte(0);
                }

                return ms.ToArray();
            }
        }

        private static byte[] ReadBytes(Stream s, int count)
        {
            byte[] buf = new byte[count];
            int read = 0;
            while (count > read)
            {
                int n = s.Read(buf, read, count - read);
                if (0 >= n) throw new EndOfStreamException();
                read += n;
            }
            return buf;
        }

        private static ushort ReadUInt16(Stream s) => (ushort)((s.ReadByte() << 8) | s.ReadByte());
        private static uint ReadUInt32(Stream s) => ((uint)s.ReadByte() << 24) | ((uint)s.ReadByte() << 16) | ((uint)s.ReadByte() << 8) | (uint)s.ReadByte();
        private static void WriteUInt16(Stream s, ushort v) { s.WriteByte((byte)(v >> 8)); s.WriteByte((byte)v); }
        private static void WriteUInt32(Stream s, uint v) { s.WriteByte((byte)(v >> 24)); s.WriteByte((byte)(v >> 16)); s.WriteByte((byte)(v >> 8)); s.WriteByte((byte)v); }

        private static ushort ReadUInt16BE(byte[] b, int o) => (ushort)((b[o] << 8) | b[o + 1]);
        private static short ReadInt16BE(byte[] b, int o) => (short)ReadUInt16BE(b, o);
        private static uint ReadUInt32BE(byte[] b, int o) => ((uint)b[o] << 24) | ((uint)b[o + 1] << 16) | ((uint)b[o + 2] << 8) | b[o + 3];
        private static void WriteUInt32BE(byte[] b, int o, uint v)
        {
            b[o] = (byte)(v >> 24);
            b[o + 1] = (byte)(v >> 16);
            b[o + 2] = (byte)(v >> 8);
            b[o + 3] = (byte)v;
        }
    }
}
