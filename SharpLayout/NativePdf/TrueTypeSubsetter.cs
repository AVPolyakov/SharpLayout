using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharpLayout.NativePdf
{
    public static class TrueTypeSubsetter
    {
        internal static readonly string[] StaticTableTags = { "OS/2", "cvt ", "fpgm", "prep" };

        public static (byte[] SubsetTtf, Dictionary<ushort, ushort> OldToNewGid) CreateSubset(
            ParsedTrueTypeFont font,
            IEnumerable<(int CodePoint, ushort GlyphId)> usedCharacters)
        {
            var charList = usedCharacters.ToList();

            // 1. Собираем все нужные глифы (включая компоненты составных глифов вроде 'ё' и 'й')
            var keepGlyphs = CollectRequiredGlyphs(font, charList);

            // 2. Перенумеровываем глифы подряд: 0, 1, 2, ... newGlyphCount - 1
            var sortedOldGids = keepGlyphs.OrderBy(g => g).ToList();
            var oldToNewGid = new Dictionary<ushort, ushort>(sortedOldGids.Count);
            for (ushort newGid = 0; sortedOldGids.Count > newGid; newGid++)
            {
                oldToNewGid[sortedOldGids[newGid]] = newGid;
            }

            ushort newGlyphCount = (ushort)sortedOldGids.Count;

            // 3. Пересобираем таблицы glyf и loca только для newGlyphCount глифов
            (byte[] newGlyf, byte[] newLoca) = BuildSubsetGlyfAndLoca(font, sortedOldGids, oldToNewGid);

            // 4. Пересобираем компактную таблицу hmtx и обновляем заголовки maxp, hhea, head
            byte[] newHmtx = RebuildHmtx(font.HmtxData, font.HheaData, sortedOldGids);

            byte[] newMaxp = (byte[])font.MaxpData.Clone();
            WriteUInt16BE(newMaxp, 4, newGlyphCount);

            byte[] newHhea = (byte[])font.HheaData.Clone();
            WriteUInt16BE(newHhea, 34, newGlyphCount);

            byte[] newHead = (byte[])font.HeadData.Clone();
            WriteUInt16BE(newHead, 50, 1);

            // 5. Строим минимальную таблицу cmap (format 4) с уже новыми GID
            var remappedChars = charList
                .Where(x => oldToNewGid.ContainsKey(x.GlyphId))
                .Select(x => (x.CodePoint, oldToNewGid[x.GlyphId]));
            byte[] newCmap = BuildMinimalCmap4(remappedChars);

            // 6. Собираем итоговый список таблиц в алфавитном порядке тегов
            var finalTables = new List<KeyValuePair<string, byte[]>>(font.StaticTables.Count + 7);
            finalTables.AddRange(font.StaticTables);
            finalTables.Add(new KeyValuePair<string, byte[]>("cmap", newCmap));
            finalTables.Add(new KeyValuePair<string, byte[]>("glyf", newGlyf));
            finalTables.Add(new KeyValuePair<string, byte[]>("head", newHead));
            finalTables.Add(new KeyValuePair<string, byte[]>("hhea", newHhea));
            finalTables.Add(new KeyValuePair<string, byte[]>("hmtx", newHmtx));
            finalTables.Add(new KeyValuePair<string, byte[]>("loca", newLoca));
            finalTables.Add(new KeyValuePair<string, byte[]>("maxp", newMaxp));
            finalTables.Sort((a, b) => string.CompareOrdinal(a.Key, b.Key));

            return (BuildTtfBinary(finalTables), oldToNewGid);
        }

        private static HashSet<ushort> CollectRequiredGlyphs(
            ParsedTrueTypeFont font,
            List<(int CodePoint, ushort GlyphId)> charList)
        {
            var keepGlyphs = new HashSet<ushort>(charList.Select(x => x.GlyphId)) { 0 };
            byte[] glyfData = font.GlyfData;
            uint[] oldOffsets = font.GlyphOffsets;
            ushort numGlyphs = font.NumGlyphs;
            var queue = new Queue<ushort>(keepGlyphs);

            const ushort ARG_1_AND_2_ARE_WORDS = 0x0001;
            const ushort WE_HAVE_A_SCALE = 0x0008;
            const ushort WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
            const ushort WE_HAVE_A_TWO_BY_TWO = 0x0080;
            const ushort MORE_COMPONENTS = 0x0020;

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

            return keepGlyphs;
        }

        private static (byte[] Glyf, byte[] Loca) BuildSubsetGlyfAndLoca(
            ParsedTrueTypeFont font,
            List<ushort> sortedOldGids,
            Dictionary<ushort, ushort> oldToNewGid)
        {
            ushort newGlyphCount = (ushort)sortedOldGids.Count;
            uint[] newOffsets = new uint[newGlyphCount + 1];
            byte[] glyfData = font.GlyfData;
            uint[] oldOffsets = font.GlyphOffsets;
            ushort numGlyphs = font.NumGlyphs;

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
                byte[] newLoca = new byte[(newGlyphCount + 1) * 4];
                for (int i = 0; newGlyphCount >= i; i++)
                {
                    WriteUInt32BE(newLoca, i * 4, newOffsets[i]);
                }
                return (newGlyf.ToArray(), newLoca);
            }
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

        internal static byte[] ReadBytes(Stream s, int count)
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

        internal static ushort ReadUInt16(Stream s) => (ushort)((s.ReadByte() << 8) | s.ReadByte());
        internal static uint ReadUInt32(Stream s) => ((uint)s.ReadByte() << 24) | ((uint)s.ReadByte() << 16) | ((uint)s.ReadByte() << 8) | (uint)s.ReadByte();

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

        internal static ushort ReadUInt16BE(byte[] b, int o) => (ushort)((b[o] << 8) | b[o + 1]);
        internal static short ReadInt16BE(byte[] b, int o) => (short)ReadUInt16BE(b, o);
        internal static uint ReadUInt32BE(byte[] b, int o) => ((uint)b[o] << 24) | ((uint)b[o + 1] << 16) | ((uint)b[o + 2] << 8) | b[o + 3];

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