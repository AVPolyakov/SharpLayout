using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace SharpLayout.NativePdf
{
    public sealed class MinimalPdfFormXObject
    {
        public double MinX { get; }
        public double MinY { get; }
        public double PointWidth { get; }
        public double PointHeight { get; }
        public byte[] ContentStream { get; }
        public bool IsContentCompressed { get; }
        public string ResourcesDictionary { get; }
        public Dictionary<int, RawPdfObject> ReferencedObjects { get; }

        public MinimalPdfFormXObject(
            double minX,
            double minY,
            double pointWidth,
            double pointHeight,
            byte[] contentStream,
            bool isContentCompressed,
            string resourcesDictionary,
            Dictionary<int, RawPdfObject> referencedObjects)
        {
            MinX = minX;
            MinY = minY;
            PointWidth = pointWidth;
            PointHeight = pointHeight;
            ContentStream = contentStream;
            IsContentCompressed = isContentCompressed;
            ResourcesDictionary = resourcesDictionary;
            ReferencedObjects = referencedObjects;
        }
    }

    public sealed class RawPdfObject
    {
        public string Dictionary { get; set; }
        public byte[] StreamData { get; set; }
    }

    internal static class PdfFormXObjectParser
    {
        private static readonly Regex ObjHeaderRegex = new Regex(@"(?:^|\r|\n)\s*(\d+)\s+\d+\s+obj\b", RegexOptions.Compiled);
        private static readonly Regex RefRegex = new Regex(@"\b(\d+)\s+0\s+R\b", RegexOptions.Compiled);

        public static MinimalPdfFormXObject ParseFirstPage(byte[] pdfBytes, Func<byte[], byte[]> zlibCompressor)
        {
            var objects = ExtractAllObjects(pdfBytes);

            // Ищем первый объект страницы (/Type /Page, но не /Type /Pages)
            RawPdfObject pageObj = null;
            foreach (var kv in objects)
            {
                if (Regex.IsMatch(kv.Value.Dictionary, @"/Type\s*/Page\b(?!s)"))
                {
                    pageObj = kv.Value;
                    break;
                }
            }

            if (pageObj == null)
                throw new InvalidDataException("No /Page object found in embedded PDF.");

            // Читаем /MediaBox (может быть на самой странице или унаследован от /Parent)
            double[] mediaBox = ParseBox(pageObj.Dictionary, "/MediaBox")
                ?? FindInheritedBox(pageObj.Dictionary, objects, "/MediaBox")
                ?? new double[] { 0, 0, 595, 842 };

            double minX = mediaBox[0];
            double minY = mediaBox[1];
            double width = mediaBox[2] - mediaBox[0];
            double height = mediaBox[3] - mediaBox[1];

            // Извлекаем /Resources (прямой словарь или ссылка N 0 R, в т.ч. унаследованная)
            string resourcesDict = ExtractResources(pageObj.Dictionary, objects);

            // Собираем поток содержимого /Contents (может быть ссылкой N 0 R или массивом [N 0 R M 0 R])
            byte[] contentBytes = ExtractPageContent(pageObj.Dictionary, objects, out bool isCompressed);
            if (!isCompressed)
            {
                contentBytes = zlibCompressor(contentBytes);
                isCompressed = true;
            }

            // Собираем транзитивно все объекты, на которые ссылается словарь /Resources
            var neededObjects = new Dictionary<int, RawPdfObject>();
            CollectReferencedObjects(resourcesDict, objects, neededObjects);

            return new MinimalPdfFormXObject(
                minX, minY, width, height,
                contentBytes, isCompressed,
                resourcesDict, neededObjects);
        }

        private static Dictionary<int, RawPdfObject> ExtractAllObjects(byte[] bytes)
        {
            var result = new Dictionary<int, RawPdfObject>();
            string text = Encoding.GetEncoding("ISO-8859-1").GetString(bytes);

            var matches = ObjHeaderRegex.Matches(text);
            for (int i = 0; i < matches.Count; i++)
            {
                int objNum = int.Parse(matches[i].Groups[1].Value, CultureInfo.InvariantCulture);
                int bodyStart = matches[i].Index + matches[i].Length;
                int endObjIdx = text.IndexOf("endobj", bodyStart, StringComparison.Ordinal);
                if (endObjIdx < 0)
                    continue;

                string bodyText = text.Substring(bodyStart, endObjIdx - bodyStart);
                int streamKeywordIdx = FindStreamKeyword(bodyText);

                if (streamKeywordIdx >= 0)
                {
                    string dictPart = bodyText.Substring(0, streamKeywordIdx).Trim();
                    int streamDataStart = bodyStart + streamKeywordIdx + 6;
                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == '\r') streamDataStart++;
                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == '\n') streamDataStart++;

                    int endStreamIdx = text.IndexOf("endstream", streamDataStart, StringComparison.Ordinal);
                    if (endStreamIdx > streamDataStart)
                    {
                        int streamLen = endStreamIdx - streamDataStart;
                        if (streamLen > 0 && bytes[streamDataStart + streamLen - 1] == '\n') streamLen--;
                        if (streamLen > 0 && bytes[streamDataStart + streamLen - 1] == '\r') streamLen--;

                        // Если в словаре указана точная числовая /Length, используем её
                        var lenMatch = Regex.Match(dictPart, @"/Length\s+(\d+)\b(?!\s+\d+\s+R)");
                        if (lenMatch.Success && int.TryParse(lenMatch.Groups[1].Value, out int declaredLen) &&
                            declaredLen >= 0 && streamDataStart + declaredLen <= endStreamIdx)
                        {
                            streamLen = declaredLen;
                        }

                        byte[] streamData = new byte[streamLen];
                        Buffer.BlockCopy(bytes, streamDataStart, streamData, 0, streamLen);
                        result[objNum] = new RawPdfObject { Dictionary = dictPart, StreamData = streamData };
                    }
                }
                else
                {
                    result[objNum] = new RawPdfObject { Dictionary = bodyText.Trim() };
                }
            }

            // Распаковываем сжатые потоки объектов (/Type /ObjStm), если PDF версии 1.5+
            var objStms = new List<RawPdfObject>();
            foreach (var obj in result.Values)
            {
                if (obj.StreamData != null && Regex.IsMatch(obj.Dictionary, @"/Type\s*/ObjStm\b"))
                    objStms.Add(obj);
            }

            foreach (var stm in objStms)
            {
                UnpackObjectStream(stm, result);
            }

            return result;
        }

        private static void UnpackObjectStream(RawPdfObject stm, Dictionary<int, RawPdfObject> target)
        {
            var nMatch = Regex.Match(stm.Dictionary, @"/N\s+(\d+)");
            var firstMatch = Regex.Match(stm.Dictionary, @"/First\s+(\d+)");
            if (!nMatch.Success || !firstMatch.Success)
                return;

            int count = int.Parse(nMatch.Groups[1].Value, CultureInfo.InvariantCulture);
            int firstOffset = int.Parse(firstMatch.Groups[1].Value, CultureInfo.InvariantCulture);

            byte[] raw = stm.Dictionary.Contains("/FlateDecode") ? DecompressZlib(stm.StreamData) : stm.StreamData;
            string unpackedText = Encoding.GetEncoding("ISO-8859-1").GetString(raw);

            string header = unpackedText.Substring(0, Math.Min(firstOffset, unpackedText.Length));
            var tokens = Regex.Matches(header, @"\d+");
            if (tokens.Count < count * 2)
                return;

            var entries = new List<(int ObjNum, int Offset)>(count);
            for (int i = 0; i < count; i++)
            {
                int num = int.Parse(tokens[i * 2].Value, CultureInfo.InvariantCulture);
                int off = int.Parse(tokens[i * 2 + 1].Value, CultureInfo.InvariantCulture);
                entries.Add((num, firstOffset + off));
            }

            for (int i = 0; i < entries.Count; i++)
            {
                int start = entries[i].Offset;
                int end = (i + 1 < entries.Count) ? entries[i + 1].Offset : unpackedText.Length;
                if (start < end && end <= unpackedText.Length)
                {
                    target[entries[i].ObjNum] = new RawPdfObject
                    {
                        Dictionary = unpackedText.Substring(start, end - start).Trim()
                    };
                }
            }
        }

        private static int FindStreamKeyword(string body)
        {
            int idx = 0;
            while ((idx = body.IndexOf("stream", idx, StringComparison.Ordinal)) >= 0)
            {
                bool validStart = idx == 0 || char.IsWhiteSpace(body[idx - 1]) || body[idx - 1] == '>';
                int after = idx + 6;
                bool validEnd = after < body.Length && (body[after] == '\r' || body[after] == '\n');
                if (validStart && validEnd)
                    return idx;
                idx += 6;
            }
            return -1;
        }

        private static double[] ParseBox(string dict, string key)
        {
            var m = Regex.Match(dict, Regex.Escape(key) + @"\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]");
            if (!m.Success)
                return null;

            return new[]
            {
                double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture),
                double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture),
                double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture),
                double.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture)
            };
        }

        private static double[] FindInheritedBox(string dict, Dictionary<int, RawPdfObject> objects, string key)
        {
            var parentMatch = Regex.Match(dict, @"/Parent\s+(\d+)\s+0\s+R");
            while (parentMatch.Success)
            {
                int parentId = int.Parse(parentMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                if (!objects.TryGetValue(parentId, out var parentObj))
                    break;
                var box = ParseBox(parentObj.Dictionary, key);
                if (box != null)
                    return box;
                parentMatch = Regex.Match(parentObj.Dictionary, @"/Parent\s+(\d+)\s+0\s+R");
            }
            return null;
        }

        private static string ExtractResources(string pageDict, Dictionary<int, RawPdfObject> objects)
        {
            string currentDict = pageDict;
            while (!string.IsNullOrEmpty(currentDict))
            {
                int resIdx = currentDict.IndexOf("/Resources", StringComparison.Ordinal);
                if (resIdx >= 0)
                {
                    string after = currentDict.Substring(resIdx + 10).TrimStart();
                    var refMatch = Regex.Match(after, @"^(\d+)\s+0\s+R\b");
                    if (refMatch.Success)
                    {
                        int resId = int.Parse(refMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                        return objects.TryGetValue(resId, out var resObj) ? resObj.Dictionary : "<< >>";
                    }
                    if (after.StartsWith("<<", StringComparison.Ordinal))
                        return ExtractBalancedDictionary(after);
                }
                var parentMatch = Regex.Match(currentDict, @"/Parent\s+(\d+)\s+0\s+R");
                if (!parentMatch.Success)
                    break;
                int parentId = int.Parse(parentMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                currentDict = objects.TryGetValue(parentId, out var pObj) ? pObj.Dictionary : null;
            }
            return "<< >>";
        }

        private static string ExtractBalancedDictionary(string text)
        {
            int depth = 0;
            for (int i = 0; i < text.Length - 1; i++)
            {
                if (text[i] == '<' && text[i + 1] == '<')
                {
                    depth++;
                    i++;
                }
                else if (text[i] == '>' && text[i + 1] == '>')
                {
                    depth--;
                    i++;
                    if (depth == 0)
                        return text.Substring(0, i + 1);
                }
            }
            return text;
        }

        private static byte[] ExtractPageContent(string pageDict, Dictionary<int, RawPdfObject> objects, out bool isCompressed)
        {
            isCompressed = false;
            var singleRef = Regex.Match(pageDict, @"/Contents\s+(\d+)\s+0\s+R\b");
            if (singleRef.Success)
            {
                int id = int.Parse(singleRef.Groups[1].Value, CultureInfo.InvariantCulture);
                var obj = objects[id];
                if (obj.Dictionary.Contains("/FlateDecode") && !obj.Dictionary.Contains("/DecodeParms"))
                {
                    isCompressed = true;
                    return obj.StreamData;
                }
                return obj.Dictionary.Contains("/FlateDecode") ? DecompressZlib(obj.StreamData) : obj.StreamData;
            }
            var arrayMatch = Regex.Match(pageDict, @"/Contents\s*\x5B([^\x5D]+)\x5D");
            if (arrayMatch.Success)
            {
                using var ms = new MemoryStream();
                foreach (Match m in RefRegex.Matches(arrayMatch.Groups[1].Value))
                {
                    int id = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                    if (objects.TryGetValue(id, out var obj) && obj.StreamData != null)
                    {
                        byte[] part = obj.Dictionary.Contains("/FlateDecode") ? DecompressZlib(obj.StreamData) : obj.StreamData;
                        ms.Write(part, 0, part.Length);
                        ms.WriteByte((byte)'\n');
                    }
                }
                return ms.ToArray();
            }
            return Array.Empty<byte>();
        }

        private static void CollectReferencedObjects(
            string dictText,
            Dictionary<int, RawPdfObject> allObjects,
            Dictionary<int, RawPdfObject> collected)
        {
            foreach (Match m in RefRegex.Matches(dictText))
            {
                int id = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                if (!collected.ContainsKey(id) && allObjects.TryGetValue(id, out var obj))
                {
                    collected[id] = obj;
                    CollectReferencedObjects(obj.Dictionary, allObjects, collected);
                }
            }
        }

        private static byte[] DecompressZlib(byte[] compressed)
        {
            int offset = (compressed.Length > 2 && compressed[0] == 0x78) ? 2 : 0;
            using var ms = new MemoryStream(compressed, offset, compressed.Length - offset);
            using var ds = new DeflateStream(ms, CompressionMode.Decompress);
            using var outMs = new MemoryStream();
            ds.CopyTo(outMs);
            return outMs.ToArray();
        }
    }
}