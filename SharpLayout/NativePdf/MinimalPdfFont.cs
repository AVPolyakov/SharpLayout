using System.Collections.Generic;

namespace SharpLayout.NativePdf;

public class MinimalPdfFont
{
    public TrueTypeMetrics Metrics { get; }
    internal string Alias { get; set; }
    internal HashSet<int> UsedCodePoints { get; } = new HashSet<int>();
    internal Dictionary<ushort, ushort> OldToNewGid { get; set; }

    public MinimalPdfFont(TrueTypeMetrics metrics)
    {
        Metrics = metrics;
    }

    internal ushort GetFinalGlyphId(int codePoint)
    {
        ushort oldGid = Metrics.GetGlyphId(codePoint);
        if (OldToNewGid != null && OldToNewGid.TryGetValue(oldGid, out ushort newGid))
            return newGid;
        return oldGid;
    }
}