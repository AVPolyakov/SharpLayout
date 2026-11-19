using System.Collections.Generic;

namespace SharpLayout.NativePdf;

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