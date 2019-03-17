using System;
using System.IO;
using PdfSharp.Fonts;
using SharpLayout.Tests.Fonts.TimesNewRoman;

namespace SharpLayout.Tests
{
    public class FontResolver : IFontResolver
    {
        private static class FamilyNames
        {
            public const string TimesNewRoman = "Times New Roman";
        }

        private static class FaceNames
        {            
            public const string TimesNewRomanBoldItalic = "TimesNewRomanBoldItalic";
            public const string TimesNewRomanBold = "TimesNewRomanBold";
            public const string TimesNewRomanItalic = "TimesNewRomanItalic";
            public const string TimesNewRoman = "TimesNewRoman";
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
                default:
                    throw new Exception("Font not found");
            }            
        }

        public byte[] GetFont(string faceName)
        {
            string fileName;
            switch (faceName)
            {
                case FaceNames.TimesNewRomanBoldItalic:
                    fileName = "timesbi.ttf";
                    break;
                case FaceNames.TimesNewRomanBold:
                    fileName = "timesbd.ttf";
                    break;
                case FaceNames.TimesNewRomanItalic:
                    fileName = "timesi.ttf";
                    break;
                case FaceNames.TimesNewRoman:
                    fileName = "times.ttf";
                    break;
                default:
                    throw new Exception("Font file not found");
            }
            var anchorType = typeof(TimesNewRoman);
            using (var stream = anchorType.Assembly.GetManifestResourceStream($"{anchorType.Namespace}.{fileName}"))
            using (var memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}