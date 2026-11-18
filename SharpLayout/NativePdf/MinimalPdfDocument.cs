using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace SharpLayout.NativePdf
{
    public class MinimalPdfDocument
    {
        private const string OpenDict = "\x3C\x3C";
        private const string CloseDict = "\x3E\x3E";
        private const string OpenArr = "\x5B";
        private const string CloseArr = "\x5D";

        private readonly List<MinimalPdfPage> _pages = new List<MinimalPdfPage>();
        private readonly Dictionary<FontKey, MinimalPdfFont> _fonts = new();

        public MinimalPdfFont GetOrRegisterFont(Font font)
        {
            var key = font.Key;
            
            if (_fonts.TryGetValue(key, out var value))
                return value;

            var trueTypeMetrics = FontStorage.GetMetrics(font);
            var minimalPdfFont = new MinimalPdfFont(trueTypeMetrics.TtfBytes) { Alias = $"F{_fonts.Count + 1}" };
            _fonts.Add(key, minimalPdfFont);
            return minimalPdfFont;
        }

        public MinimalPdfPage AddPage(double width = 595, double height = 842)
        {
            var page = new MinimalPdfPage(width, height);
            _pages.Add(page);
            return page;
        }

        public void Save(Stream output)
        {
            var objects = new List<byte[]>();
            int AddObject(byte[] data) { objects.Add(data); return objects.Count; }
            int ReserveObject() { objects.Add(null); return objects.Count; }

            int catalogId = ReserveObject();
            int pagesId = ReserveObject();

            var fontResources = new StringBuilder($"{OpenDict} ");
            foreach (var font in _fonts.Values)
            {
                var m = font.Metrics;
                var usedGids = font.UsedCodePoints.Select(cp => m.GetGlyphId(cp));
                byte[] subsetTtf = TrueTypeSubsetter.CreateSubset(font.RawBytes, usedGids);
                byte[] compressedTtf = CompressZlib(subsetTtf);
                int ttfStreamId = AddObject(BuildStreamObject(
                    $"{OpenDict} /Length {compressedTtf.Length} /Length1 {subsetTtf.Length} /Filter /FlateDecode {CloseDict}",
                    compressedTtf));

                string familyName = (!string.IsNullOrEmpty(m.FontFamily) ? m.FontFamily : $"Font_{font.Alias}").Replace(" ", "#20");
                string styleSuffix = "";
                if (m.IsBold && m.IsItalic) styleSuffix = ",BoldItalic";
                else if (m.IsBold) styleSuffix = ",Bold";
                else if (m.IsItalic) styleSuffix = ",Italic";
                string pdfFontName = $"SHARPL+{familyName}{styleSuffix}";
                
                int ascent = m.WinAscent > 0 ? m.WinAscent : m.Ascender;
                int descent = m.WinDescent > 0 ? -m.WinDescent : m.Descender;

                int descriptorId = AddObject(Encoding.ASCII.GetBytes(
                    $"{OpenDict} /Type /FontDescriptor /FontName /{pdfFontName} /Flags 4 " +
                    $"/FontBBox {OpenArr}-500 {descent} 1500 {ascent}{CloseArr} " +
                    $"/ItalicAngle 0 /Ascent {ascent} /Descent {descent} /CapHeight {ascent} /StemV 80 " +
                    $"/FontFile2 {ttfStreamId} 0 R {CloseDict}"));

                var widths = new StringBuilder($"{OpenArr} ");
                var sortedGids = new SortedSet<ushort>();
                foreach (int cp in font.UsedCodePoints)
                {
                    ushort gid = m.GetGlyphId(cp);
                    if (gid != 0) sortedGids.Add(gid);
                }

                foreach (ushort gid in sortedGids)
                {
                    int w = (int)Math.Round((double)m.GetAdvanceWidth(gid) / m.UnitsPerEm * 1000.0);
                    widths.Append($"{gid} {OpenArr}{w}{CloseArr} ");
                }
                widths.Append(CloseArr);

                int cidFontId = AddObject(Encoding.ASCII.GetBytes(
                    $"{OpenDict} /Type /Font /Subtype /CIDFontType2 /BaseFont /{pdfFontName} " +
                    $"/CIDSystemInfo {OpenDict} /Registry (Adobe) /Ordering (Identity) /Supplement 0 {CloseDict} " +
                    $"/FontDescriptor {descriptorId} 0 R /DW 1000 /W {widths} /CIDToGIDMap /Identity {CloseDict}"));

                byte[] cmapBytes = CompressZlib(BuildToUnicodeCMap(font));
                int toUnicodeId = AddObject(BuildStreamObject(
                    $"{OpenDict} /Length {cmapBytes.Length} /Filter /FlateDecode {CloseDict}", cmapBytes));

                int type0Id = AddObject(Encoding.ASCII.GetBytes(
                    $"{OpenDict} /Type /Font /Subtype /Type0 /BaseFont /{pdfFontName}-Identity-H " +
                    $"/Encoding /Identity-H /DescendantFonts {OpenArr}{cidFontId} 0 R{CloseArr} /ToUnicode {toUnicodeId} 0 R {CloseDict}"));

                fontResources.Append($"/{font.Alias} {type0Id} 0 R ");
            }
            fontResources.Append(CloseDict);

            var pageIds = new List<int>();
            foreach (var page in _pages)
            {
                byte[] rawContent = Encoding.ASCII.GetBytes(page.GetContent());
                byte[] compressedContent = CompressZlib(rawContent);
                int contentId = AddObject(BuildStreamObject(
                    $"{OpenDict} /Length {compressedContent.Length} /Filter /FlateDecode {CloseDict}", compressedContent));

                int pageId = AddObject(Encoding.ASCII.GetBytes(string.Format(CultureInfo.InvariantCulture,
                    "{0} /Type /Page /Parent {1} 0 R /MediaBox {2}0 0 {3:0.###} {4:0.###}{5} " +
                    "/Resources {0} /ProcSet {2}/PDF /Text{5} /Font {6} {7} /Contents {8} 0 R {7}",
                    OpenDict, pagesId, OpenArr, page.Width, page.Height, CloseArr, fontResources, CloseDict, contentId)));

                pageIds.Add(pageId);
            }

            string kids = string.Join(" ", pageIds.Select(id => $"{id} 0 R"));
            byte[] pagesBytes = Encoding.ASCII.GetBytes($"{OpenDict} /Type /Pages /Kids {OpenArr}{kids}{CloseArr} /Count {pageIds.Count} {CloseDict}");
            objects.RemoveAt(pagesId - 1);
            objects.Insert(pagesId - 1, pagesBytes);

            byte[] catalogBytes = Encoding.ASCII.GetBytes($"{OpenDict} /Type /Catalog /Pages {pagesId} 0 R {CloseDict}");
            objects.RemoveAt(catalogId - 1);
            objects.Insert(catalogId - 1, catalogBytes);

            using (var ms = new MemoryStream())
            {
                WriteAscii(ms, "%PDF-1.7\n%");
                ms.Write(new byte[] { 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }, 0, 5);

                var offsets = new List<long> { 0 };
                for (int i = 0; objects.Count > i; i++)
                {
                    byte[] objBytes = objects.ElementAt(i);
                    offsets.Add(ms.Position);
                    WriteAscii(ms, $"{i + 1} 0 obj\n");
                    ms.Write(objBytes, 0, objBytes.Length);
                    WriteAscii(ms, "\nendobj\n");
                }

                long xrefPos = ms.Position;
                WriteAscii(ms, $"xref\n0 {objects.Count + 1}\n");
                WriteAscii(ms, "0000000000 65535 f \n");
                for (int i = 1; objects.Count >= i; i++)
                {
                    long offset = offsets.ElementAt(i);
                    WriteAscii(ms, $"{offset:D10} 00000 n \n");
                }

                WriteAscii(ms, $"trailer\n{OpenDict} /Size {objects.Count + 1} /Root {catalogId} 0 R {CloseDict}\nstartxref\n{xrefPos}\n%%EOF");
                ms.Position = 0;
                ms.CopyTo(output);
            }
        }

        private static byte[] BuildToUnicodeCMap(MinimalPdfFont font)
        {
            var sb = new StringBuilder();
            sb.Append("/CIDInit /ProcSet findresource begin\n");
            sb.Append("12 dict begin\n");
            sb.Append("begincmap\n");
            sb.Append($"/CIDSystemInfo {OpenDict} /Registry (Adobe) /Ordering (UCS) /Supplement 0 {CloseDict} def\n");
            sb.Append("/CMapName /Adobe-Identity-UCS def\n");
            sb.Append("/CMapType 2 def\n");
            sb.Append("1 begincodespacerange\n");
            sb.Append("<0000> <FFFF>\n");
            sb.Append("endcodespacerange\n");

            var mappings = new SortedDictionary<ushort, int>();
            foreach (int cp in font.UsedCodePoints)
            {
                ushort gid = font.Metrics.GetGlyphId(cp);
                if (gid != 0) mappings[gid] = cp;
            }

            if (mappings.Count > 0)
            {
                sb.Append($"{mappings.Count} beginbfchar\n");
                foreach (var kv in mappings)
                {
                    if (kv.Value <= 0xFFFF)
                    {
                        sb.AppendFormat("<{0:X4}> <{1:X4}>\n", kv.Key, kv.Value);
                    }
                    else
                    {
                        string surrogate = char.ConvertFromUtf32(kv.Value);
                        sb.AppendFormat("<{0:X4}> <{1:X4}{2:X4}>\n", kv.Key, (int)surrogate[0], (int)surrogate[1]);
                    }
                }
                sb.Append("endbfchar\n");
            }

            sb.Append("endcmap\n");
            sb.Append("CMapName currentdict /CMap defineresource pop\n");
            sb.Append("end\n");
            sb.Append("end\n");
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static byte[] BuildStreamObject(string dict, byte[] data)
        {
            var head = Encoding.ASCII.GetBytes(dict + "\nstream\n");
            var tail = Encoding.ASCII.GetBytes("\nendstream");
            var res = new byte[head.Length + data.Length + tail.Length];
            Buffer.BlockCopy(head, 0, res, 0, head.Length);
            Buffer.BlockCopy(data, 0, res, head.Length, data.Length);
            Buffer.BlockCopy(tail, 0, res, head.Length + data.Length, tail.Length);
            return res;
        }

        private static byte[] CompressZlib(byte[] data)
        {
            using (var outMs = new MemoryStream())
            {
                outMs.WriteByte(0x78);
                outMs.WriteByte(0x9C);
                using (var ds = new DeflateStream(outMs, CompressionLevel.Optimal, true))
                    ds.Write(data, 0, data.Length);
                uint a = 1, b = 0;
                foreach (byte d in data) { a = (a + d) % 65521; b = (b + a) % 65521; }
                uint adler = (b << 16) | a;
                outMs.WriteByte((byte)(adler >> 24));
                outMs.WriteByte((byte)(adler >> 16));
                outMs.WriteByte((byte)(adler >> 8));
                outMs.WriteByte((byte)adler);
                return outMs.ToArray();
            }
        }
        private static void WriteAscii(Stream s, string text)
        {
            var b = Encoding.ASCII.GetBytes(text);
            s.Write(b, 0, b.Length);
        }
    }
}