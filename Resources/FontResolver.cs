using System;
using System.IO;
using PdfSharp.Fonts;
using Resources.Fonts.Arial;
using Resources.Fonts.TimesNewRoman;
using Resources.Fonts.Wingdings;
using Resources.Fonts.Wingdings2;

namespace Resources
{
    public class FontResolver : IFontResolver
    {
        private static class FamilyNames
        {
            public const string TimesNewRoman = "Times New Roman";
            public const string Wingdings = "Wingdings";
            public const string Wingdings2 = "Wingdings 2";
            public const string Arial = "Arial";
        }

        private static class FaceNames
        {            
            public const string TimesNewRomanBoldItalic = "TimesNewRomanBoldItalic";
            public const string TimesNewRomanBold = "TimesNewRomanBold";
            public const string TimesNewRomanItalic = "TimesNewRomanItalic";
            public const string TimesNewRoman = "TimesNewRoman";
            public const string Wingdings = "Wingdings";
            public const string Wingdings2 = "Wingdings2";
            public const string ArialBoldItalic = "ArialBoldItalic";
            public const string ArialBold = "ArialBold";
            public const string ArialItalic = "ArialItalic";
            public const string Arial = "Arial";
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            switch (familyName)
            {
                case FamilyNames.TimesNewRoman:
                    if (isBold)
                    {
                        if (isItalic)
                            return new FontResolverInfo(FaceNames.TimesNewRomanBoldItalic);
                        else
                            return new FontResolverInfo(FaceNames.TimesNewRomanBold);
                    }
                    else
                    {
                        if (isItalic)
                            return new FontResolverInfo(FaceNames.TimesNewRomanItalic);
                        else
                            return new FontResolverInfo(FaceNames.TimesNewRoman);
                    }
                case FamilyNames.Wingdings:
                    return new FontResolverInfo(FaceNames.Wingdings, isBold, isItalic);
                case FamilyNames.Wingdings2:
                    return new FontResolverInfo(FaceNames.Wingdings2, isBold, isItalic);
                case FamilyNames.Arial:
                    if (isBold)
                    {
                        if (isItalic)
                            return new FontResolverInfo(FaceNames.ArialBoldItalic);
                        else
                            return new FontResolverInfo(FaceNames.ArialBold);
                    }
                    else
                    {
                        if (isItalic)
                            return new FontResolverInfo(FaceNames.ArialItalic);
                        else
                            return new FontResolverInfo(FaceNames.Arial);
                    }
            }       
            return new FontResolverInfo(FaceNames.Arial);
        }

        public byte[] GetFont(string faceName)
        {
            Type anchorType;
            string fileName;
            switch (faceName)
            {
                case FaceNames.TimesNewRomanBoldItalic:
                    anchorType = typeof(TimesNewRoman);
                    fileName = "timesbi.ttf";
                    break;
                case FaceNames.TimesNewRomanBold:
                    anchorType = typeof(TimesNewRoman);
                    fileName = "timesbd.ttf";
                    break;
                case FaceNames.TimesNewRomanItalic:
                    anchorType = typeof(TimesNewRoman);
                    fileName = "timesi.ttf";
                    break;
                case FaceNames.TimesNewRoman:
                    anchorType = typeof(TimesNewRoman);
                    fileName = "times.ttf";
                    break;
                case FaceNames.Wingdings:
                    anchorType = typeof(Wingdings);
                    fileName = "wingding.ttf";
                    break;
                case FaceNames.Wingdings2:
                    anchorType = typeof(Wingdings2);
                    fileName = "WINGDNG2.TTF";
                    break;
                case FaceNames.ArialBoldItalic:
                    anchorType = typeof(Arial);
                    fileName = "arialbi.ttf";
                    break;
                case FaceNames.ArialBold:
                    anchorType = typeof(Arial);
                    fileName = "arialbd.ttf";
                    break;
                case FaceNames.ArialItalic:
                    anchorType = typeof(Arial);
                    fileName = "ariali.ttf";
                    break;
                case FaceNames.Arial:
                    anchorType = typeof(Arial);
                    fileName = "arial.ttf";
                    break;
                default:
                    throw new Exception($"Font file not found, FaceName = {faceName}");
            }
            using (var stream = anchorType.Assembly.GetManifestResourceStream($"{anchorType.Namespace}.{fileName}"))
            using (var memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}