using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharpLayout.NativePdf
{
    public static class TrueTypeSubsetter
    {
        public static (byte[] SubsetTtf, Dictionary<ushort, ushort> OldToNewGid) CreateSubset(
            byte[] ttfBytes,
            IEnumerable<(int CodePoint, ushort GlyphId)> usedCharacters)
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

            var charList = usedCharacters.ToList();

            // 3. Собираем все нужные глифы (включая компоненты составных глифов вроде 'ё' и 'й')
            var keepGlyphs = new HashSet<ushort>(charList.Select(x => x.GlyphId)) { 0 };
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

            // 4. Перенумеровываем глифы подряд: 0, 1, 2, ... newGlyphCount - 1
            var sortedOldGids = keepGlyphs.OrderBy(g => g).ToList();
            var oldToNewGid = new Dictionary<ushort, ushort>(sortedOldGids.Count);
            for (ushort newGid = 0; sortedOldGids.Count > newGid; newGid++)
            {
                oldToNewGid[sortedOldGids[newGid]] = newGid;
            }

            ushort newGlyphCount = (ushort)sortedOldGids.Count;

            // 5. Пересобираем таблицы glyf и loca только для newGlyphCount глифов с обновлением составных ссылок
            uint[] newOffsets = new uint[newGlyphCount + 1];
            using (var newGlyf = new MemoryStream())
            {
                for (ushort newGid = 0; newGlyphCount > newGid; newGid++)
                {
                    newOffsets[newGid] = (uint)newGlyf.Position;
                    ushort oldGid = sortedOldGids[newGid];
                    if (oldGid >= numGlyphs) continue;

                    uint start = oldOffsets[oldGid];
                    uint length = oldOffsets[oldGid + 1] - start;
                    if (length > 0 && start + length <= glyfData.Length)
                    {
                        byte[] glyphChunk = new byte[length];
                        Buffer.BlockCopy(glyfData, (int)start, glyphChunk, 0, (int)length);

                        // Если составной глиф — обновляем индексы компонентов на новые GID
                        short numberOfContours = ReadInt16BE(glyphChunk, 0);
                        if (0 > numberOfContours)
                        {
                            RemapCompositeGlyphComponents(glyphChunk, oldToNewGid);
                        }

                        newGlyf.Write(glyphChunk, 0, glyphChunk.Length);
                        while ((newGlyf.Position & 3) != 0) newGlyf.WriteByte(0);
                    }
                }
                newOffsets[newGlyphCount] = (uint)newGlyf.Position;
                tables["glyf"] = newGlyf.ToArray();
            }

            // Записываем компактную таблицу loca (32-bit)
            byte[] newLoca = new byte[(newGlyphCount + 1) * 4];
            for (int i = 0; newGlyphCount >= i; i++)
            {
                WriteUInt32BE(newLoca, i * 4, newOffsets[i]);
            }
            tables["loca"] = newLoca;

            // 6. Пересобираем компактную таблицу hmtx ровно на newGlyphCount записей
            tables["hmtx"] = RebuildHmtx(tables["hmtx"], tables["hhea"], sortedOldGids);

            // 7. Обновляем numGlyphs в maxp и numberOfHMetrics в hhea
            byte[] newMaxp = (byte[])tables["maxp"].Clone();
            WriteUInt16BE(newMaxp, 4, newGlyphCount);
            tables["maxp"] = newMaxp;

            byte[] newHhea = (byte[])tables["hhea"].Clone();
            WriteUInt16BE(newHhea, 34, newGlyphCount);
            tables["hhea"] = newHhea;

            // Указываем в head, что формат loca теперь 1 (32-bit)
            byte[] newHead = (byte[])tables["head"].Clone();
            newHead[50] = 0;
            newHead[51] = 1;
            tables["head"] = newHead;

            // 8. Строим минимальную таблицу cmap (format 4) с уже новыми GID
            var remappedChars = charList
                .Where(x => oldToNewGid.ContainsKey(x.GlyphId))
                .Select(x => (x.CodePoint, oldToNewGid[x.GlyphId]));
            tables["cmap"] = BuildMinimalCmap4(remappedChars);

            string[] requiredTags = { "head", "hhea", "maxp", "OS/2", "hmtx", "cmap", "loca", "glyf", "cvt ", "fpgm", "prep" };
            var finalTables = tables
                .Where(kv => requiredTags.Contains(kv.Key))
                .OrderBy(kv => kv.Key, StringComparer.Ordinal)
                .ToList();

            return (BuildTtfBinary(finalTables), oldToNewGid);
        }

        private static void RemapCompositeGlyphComponents(byte[] glyphChunk, Dictionary<ushort, ushort> oldToNewGid)
        {
            int pos = 10;
            const ushort ARG_1_AND_2_ARE_WORDS = 0x0001;
            const ushort WE_HAVE_A_SCALE = 0x0008;
            const ushort WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
            const ushort WE_HAVE_A_TWO_BY_TWO = 0x0080;
            const ushort MORE_COMPONENTS = 0x0020;

            while (pos + 4 <= glyphChunk.Length)
            {
                ushort flags = ReadUInt16BE(glyphChunk, pos);
                ushort oldComponentGid = ReadUInt16BE(glyphChunk, pos + 2);
                if (oldToNewGid.TryGetValue(oldComponentGid, out ushort newComponentGid))
                {
                    WriteUInt16BE(glyphChunk, pos + 2, newComponentGid);
                }
                pos += 4;

                pos += (flags & ARG_1_AND_2_ARE_WORDS) != 0 ? 4 : 2;
                if ((flags & WE_HAVE_A_SCALE) != 0) pos += 2;
                else if ((flags & WE_HAVE_AN_X_AND_Y_SCALE) != 0) pos += 4;
                else if ((flags & WE_HAVE_A_TWO_BY_TWO) != 0) pos += 8;

                if ((flags & MORE_COMPONENTS) == 0) break;
            }
        }

        private static byte[] RebuildHmtx(byte[] oldHmtx, byte[] hhea, List<ushort> sortedOldGids)
        {
            ushort oldNumHMetrics = ReadUInt16BE(hhea, 34);
            ushort lastAdvance = oldNumHMetrics > 0 ? ReadUInt16BE(oldHmtx, (oldNumHMetrics - 1) * 4) : (ushort)0;
            int lsbArrayStart = oldNumHMetrics * 4;

            byte[] newHmtx = new byte[sortedOldGids.Count * 4];
            for (int i = 0; sortedOldGids.Count > i; i++)
            {
                ushort oldGid = sortedOldGids[i];
                ushort advance;
                short lsb;

                if (oldNumHMetrics > oldGid)
                {
                    advance = ReadUInt16BE(oldHmtx, oldGid * 4);
                    lsb = ReadInt16BE(oldHmtx, oldGid * 4 + 2);
                }
                else
                {
                    advance = lastAdvance;
                    int lsbOffset = lsbArrayStart + (oldGid - oldNumHMetrics) * 2;
                    lsb = (lsbOffset + 2 <= oldHmtx.Length) ? ReadInt16BE(oldHmtx, lsbOffset) : (short)0;
                }
                WriteUInt16BE(newHmtx, i * 4, advance);
                WriteUInt16BE(newHmtx, i * 4 + 2, (ushort)lsb);
            }
            return newHmtx;
        }

        private static byte[] BuildMinimalCmap4(IEnumerable<(int CodePoint, ushort GlyphId)> mappings)
        {
            var sorted = mappings
                .Where(x => x.CodePoint > 0 && 0xFFFF > x.CodePoint && x.GlyphId != 0)
                .GroupBy(x => x.CodePoint)
                .Select(g => (Code: (ushort)g.Key, Gid: g.First().GlyphId))
                .OrderBy(x => x.Code)
                .ToList();
            var segments = new List<(ushort Start, ushort End, short Delta)>();
            int idx = 0;
            while (sorted.Count > idx)
            {
                ushort start = sorted[idx].Code;
                ushort end = start;
                short delta = (short)(sorted[idx].Gid - start);
                idx++;
                while (sorted.Count > idx &&
                       sorted[idx].Code == end + 1 &&
                       (short)(sorted[idx].Gid - sorted[idx].Code) == delta)
                {
                    end = sorted[idx].Code;
                    idx++;
                }
                segments.Add((start, end, delta));
            }
            segments.Add((0xFFFF, 0xFFFF, 1));
            ushort segCount = (ushort)segments.Count;
            ushort segCountX2 = (ushort)(segCount * 2);
            ushort entrySelector = 0;
            ushort searchRange = 2;
            while (segCountX2 >= searchRange * 2)
            {
                searchRange *= 2;
                entrySelector++;
            }
            ushort rangeShift = (ushort)(segCountX2 - searchRange);
            ushort subtableLength = (ushort)(16 + segCount * 8);
            using (var ms = new MemoryStream())
            {
                WriteUInt16(ms, 0);
                WriteUInt16(ms, 1);
                WriteUInt16(ms, 3);
                WriteUInt16(ms, 1);
                WriteUInt32(ms, 12);
                WriteUInt16(ms, 4);
                WriteUInt16(ms, subtableLength);
                WriteUInt16(ms, 0);
                WriteUInt16(ms, segCountX2);
                WriteUInt16(ms, searchRange);
                WriteUInt16(ms, entrySelector);
                WriteUInt16(ms, rangeShift);
                foreach (var seg in segments) WriteUInt16(ms, seg.End);
                WriteUInt16(ms, 0);
                foreach (var seg in segments) WriteUInt16(ms, seg.Start);
                foreach (var seg in segments) WriteUInt16(ms, (ushort)seg.Delta);
                for (int i = 0; segCount > i; i++) WriteUInt16(ms, 0);
                return ms.ToArray();
            }
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
                WriteUInt32(ms, 0x00010000);
                WriteUInt16(ms, numTables);
                WriteUInt16(ms, searchRange);
                WriteUInt16(ms, entrySelector);
                WriteUInt16(ms, rangeShift);
                uint currentOffset = (uint)(12 + numTables * 16);
                foreach (var kv in tables)
                {
                    byte[] tagBytes = Encoding.ASCII.GetBytes(kv.Key.PadRight(4));
                    ms.Write(tagBytes, 0, 4);
                    WriteUInt32(ms, 0);
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

        private static void WriteUInt16(Stream s, ushort v)
        {
            s.WriteByte((byte)(v >> 8));
            s.WriteByte((byte)v);
        }

        private static void WriteUInt32(Stream s, uint v)
        {
            s.WriteByte((byte)(v >> 24));
            s.WriteByte((byte)(v >> 16));
            s.WriteByte((byte)(v >> 8));
            s.WriteByte((byte)v);
        }

        private static ushort ReadUInt16BE(byte[] b, int o) => (ushort)((b[o] << 8) | b[o + 1]);
        private static short ReadInt16BE(byte[] b, int o) => (short)ReadUInt16BE(b, o);
        private static uint ReadUInt32BE(byte[] b, int o) => ((uint)b[o] << 24) | ((uint)b[o + 1] << 16) | ((uint)b[o + 2] << 8) | b[o + 3];

        private static void WriteUInt16BE(byte[] b, int o, ushort v)
        {
            b[o] = (byte)(v >> 8);
            b[o + 1] = (byte)v;
        }

        private static void WriteUInt32BE(byte[] b, int o, uint v)
        {
            b[o] = (byte)(v >> 24);
            b[o + 1] = (byte)(v >> 16);
            b[o + 2] = (byte)(v >> 8);
            b[o + 3] = (byte)v;
        }
    }
}