using System.Collections.Concurrent;
using PdfSharp.Drawing;
using PdfSharp.Fonts;

namespace SharpLayout.NativePdf;

public static class FontStorage
{
    private static readonly ConcurrentDictionary<FontResolutionKey, TrueTypeMetrics> _trueTypeMetricsDictionary = new();
    
    private static TrueTypeMetrics GetMetrics(Font font)
    {
        var key = new FontResolutionKey(font.FamilyInfo.Name,
            isBold: font.Style.HasFlag(XFontStyle.Bold), isItalic: font.Style.HasFlag(XFontStyle.Italic));

        return _trueTypeMetricsDictionary.GetOrAdd(
            key,
            _ =>
            {
                var fontResolverInfo = GlobalFontSettings.FontResolver.ResolveTypeface(font.FamilyInfo.Name,
                    isBold: font.Style.HasFlag(XFontStyle.Bold), isItalic: font.Style.HasFlag(XFontStyle.Italic));
                var bytes = GlobalFontSettings.FontResolver.GetFont(fontResolverInfo.FaceName);
                
                return new TrueTypeMetrics(bytes, simulateBold: fontResolverInfo.MustSimulateBold);
            });
    }
    
    public static XSize MeasureString(string text, Font font)
    {
        var size = GetMetrics(font).MeasureString(text, font.Size);
        return new XSize(size.Width, size.Height);
    }    
}