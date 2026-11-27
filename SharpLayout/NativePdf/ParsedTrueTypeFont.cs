using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static SharpLayout.NativePdf.TrueTypeSubsetter;

namespace SharpLayout.NativePdf;

public sealed class ParsedTrueTypeFont
{
    public IReadOnlyDictionary<string, byte[]> Tables { get; }
    public ushort NumGlyphs { get; }
    public uint[] GlyphOffsets { get; }
    public byte[] GlyfData { get; }
    public byte[] HmtxData { get; }
    public byte[] HheaData { get; }
    public byte[] MaxpData { get; }
    public byte[] HeadData { get; }
    public List<KeyValuePair<string, byte[]>> StaticTables { get; }

    internal ParsedTrueTypeFont(
        Dictionary<string, byte[]> tables,
        ushort numGlyphs,
        uint[] glyphOffsets,
        List<KeyValuePair<string, byte[]>> staticTables)
    {
        Tables = tables;
        NumGlyphs = numGlyphs;
        GlyphOffsets = glyphOffsets;
        GlyfData = tables["glyf"];
        HmtxData = tables["hmtx"];
        HheaData = tables["hhea"];
        MaxpData = tables["maxp"];
        HeadData = tables["head"];
        StaticTables = staticTables;
    }

    public static ParsedTrueTypeFont ParseFontStructure(byte[] ttfBytes)
    {
        var tables = new Dictionary<string, byte[]>(StringComparer.Ordinal);

        using (var ms = new MemoryStream(ttfBytes))
        {
            ReadUInt32(ms); // sfntVersion
            ushort numTables = ReadUInt16(ms);
            ms.Position += 6;

            var entries = new List<(string Tag, uint Offset, uint Length)>(numTables);
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

        short indexToLocFormat = ReadInt16BE(tables["head"], 50);
        ushort numGlyphs = ReadUInt16BE(tables["maxp"], 4);

        uint[] offsets = new uint[numGlyphs + 1];
        byte[] locaData = tables["loca"];
        for (int i = 0; numGlyphs >= i; i++)
        {
            offsets[i] = indexToLocFormat == 0
                ? (uint)(ReadUInt16BE(locaData, i * 2) * 2)
                : ReadUInt32BE(locaData, i * 4);
        }

        var staticTables = new List<KeyValuePair<string, byte[]>>(StaticTableTags.Length);
        foreach (string tag in StaticTableTags)
        {
            if (tables.TryGetValue(tag, out byte[] data))
            {
                staticTables.Add(new KeyValuePair<string, byte[]>(tag, data));
            }
        }

        return new ParsedTrueTypeFont(tables, numGlyphs, offsets, staticTables);
    }
}