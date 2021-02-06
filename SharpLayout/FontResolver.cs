using System;
using System.Collections.Concurrent;
using System.IO;
using System.IO.Compression;
using PdfSharp.Fonts;
using SharpLayout.Fonts;
using SixLabors.Fonts;

namespace SharpLayout
{
    public class FontResolver : IFontResolver
    {
        static FontResolver() => GlobalFontSettings.FontResolver = new FontResolver();

        private FontResolver()
        {
        }

        private static readonly ConcurrentDictionary<FontResolutionKey, FontResolverInfo> fontResolverInfos = new ();
        private static readonly ConcurrentDictionary<string, IHandler<Stream>> fontFaces = new ();
        private static readonly string defaultFontFace = Guid.NewGuid().ToString();
        
        public static void Install(FontSlot slot, IHandler<Stream> streamHandler)
        {
            var fontDescription = streamHandler.Handle(FontDescription.LoadDescription);
                
            var faceName = Guid.NewGuid().ToString();
            fontFaces.TryAdd(faceName, streamHandler);
            
            var isBold = fontDescription.Style.HasFlag(FontStyle.Bold);
            var isItalic = fontDescription.Style.HasFlag(FontStyle.Italic);
            var familyName = slot.FamilyInfo(fontDescription.FontFamilyInvariantCulture).FullName;
            fontResolverInfos.TryAdd(
                new FontResolutionKey(familyName, isBold, isItalic), 
                new FontResolverInfo(faceName));
        }

        public static void Install(FontSlot slot, Type anchorType)
        {
            foreach (var resourceName in anchorType.Assembly.GetManifestResourceNames())
                if (resourceName.StartsWith(anchorType.Namespace))
                    Install(slot, new StreamHandler(resourceName, anchorType));
        }

        private class StreamHandler: IHandler<Stream>
        {
            private readonly string resourceName;
            private readonly Type anchorType;

            public StreamHandler(string resourceName, Type anchorType)
            {
                this.resourceName = resourceName;
                this.anchorType = anchorType;
            }

            public TResult Handle<TResult>(Func<Stream, TResult> func)
            {
                using (var resourceStream = anchorType.Assembly.GetManifestResourceStream(resourceName))
                {
                    var extension = Path.GetExtension(resourceName);
                    if (".gz".Equals(extension, StringComparison.InvariantCultureIgnoreCase))
                    {
                        using (var gZipStream = new GZipStream(resourceStream, CompressionMode.Decompress))
                        using (var memoryStream = new MemoryStream())
                        {
                            gZipStream.CopyTo(memoryStream);
                            memoryStream.Position = 0;
                            return func(memoryStream);
                        }
                    }
                    else
                    {
                        return func(resourceStream);
                    }
                }
            }
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (fontResolverInfos.TryGetValue(new FontResolutionKey(familyName, isBold, isItalic), out var fontResolverInfo))
                return fontResolverInfo;

            if (fontResolverInfos.TryGetValue(new FontResolutionKey(familyName, false, false), out var regularFontResolverInfo))
                return new FontResolverInfo(regularFontResolverInfo.FaceName, isBold, isItalic);

            return new FontResolverInfo(defaultFontFace, isBold, isItalic);
        }

        public byte[] GetFont(string faceName)
        {
            if (fontFaces.TryGetValue(faceName, out var streamHandler))
                return streamHandler.Handle(stream =>
                {
                    var bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, bytes.Length);
                    return bytes;
                });

            var anchorType = typeof(FontAnchor);
            var fileName = "Roboto-Regular.ttf.gz";
            using (var resourceStream = anchorType.Assembly.GetManifestResourceStream($"{anchorType.Namespace}.{fileName}"))
            using (var gZipStream = new GZipStream(resourceStream, CompressionMode.Decompress))
            using (var memoryStream = new MemoryStream())
            {
                gZipStream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }

        internal static void Init()
        {
        }
    }
}