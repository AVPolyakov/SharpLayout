using System;
using System.IO;
using System.IO.Compression;

namespace SharpLayout.NativePdf
{
    public sealed class MinimalPdfImage
    {
        internal string Alias { get; set; }
        public int PixelWidth { get; }
        public int PixelHeight { get; }
        public double PointWidth { get; }
        public double PointHeight { get; }
        public bool IsJpeg { get; }
        public int BitsPerComponent { get; }
        public string ColorSpace { get; } // "/DeviceRGB", "/DeviceGray" или "/DeviceCMYK"
        public bool InvertCmyk { get; }
        public byte[] CompressedData { get; }
        public byte[] AlphaMaskCompressedData { get; }

        public MinimalPdfImage(
            int pixelWidth,
            int pixelHeight,
            double dpiX,
            double dpiY,
            bool isJpeg,
            int bitsPerComponent,
            string colorSpace,
            byte[] compressedData,
            byte[] alphaMaskCompressedData = null,
            bool invertCmyk = false)
        {
            PixelWidth = pixelWidth;
            PixelHeight = pixelHeight;
            PointWidth = pixelWidth * 72.0 / (dpiX > 0 ? dpiX : 96.0);
            PointHeight = pixelHeight * 72.0 / (dpiY > 0 ? dpiY : 96.0);
            IsJpeg = isJpeg;
            BitsPerComponent = bitsPerComponent;
            ColorSpace = colorSpace;
            InvertCmyk = invertCmyk;
            CompressedData = compressedData;
            AlphaMaskCompressedData = alphaMaskCompressedData;
        }
    }

    internal static class PdfImageParser
    {
        public static MinimalPdfImage Parse(byte[] bytes, Func<byte[], byte[]> zlibCompressor)
        {
            if (bytes == null || bytes.Length < 8)
                throw new InvalidDataException("Image data is empty or corrupted.");

            // JPEG magic: FF D8
            if (bytes[0] == 0xFF && bytes[1] == 0xD8)
                return ParseJpeg(bytes);

            // PNG magic: 89 50 4E 47 0D 0A 1A 0A
            if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                return ParsePng(bytes, zlibCompressor);

            throw new NotSupportedException("Only JPEG and PNG image formats are supported.");
        }

        private static MinimalPdfImage ParseJpeg(byte[] bytes)
        {
            int pos = 2;
            double dpiX = 96.0;
            double dpiY = 96.0;
            bool hasAdobeApp14 = false;

            while (pos + 4 <= bytes.Length)
            {
                if (bytes[pos] != 0xFF)
                {
                    pos++;
                    continue;
                }

                byte marker = bytes[pos + 1];
                pos += 2;

                if (marker == 0xD9 || marker == 0xDA) // EOI или SOS
                    break;

                if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7)) // Standalone markers
                    continue;

                if (pos + 2 > bytes.Length)
                    break;

                int segmentLength = (bytes[pos] << 8) | bytes[pos + 1];
                if (segmentLength < 2 || pos + segmentLength > bytes.Length)
                    break;

                // APP0 JFIF (читаем DPI)
                if (marker == 0xE0 && segmentLength >= 16 &&
                    bytes[pos + 2] == 'J' && bytes[pos + 3] == 'F' &&
                    bytes[pos + 4] == 'I' && bytes[pos + 5] == 'F' && bytes[pos + 6] == 0)
                {
                    byte units = bytes[pos + 9];
                    int xDensity = (bytes[pos + 10] << 8) | bytes[pos + 11];
                    int yDensity = (bytes[pos + 12] << 8) | bytes[pos + 13];
                    if (units == 1 && xDensity > 0 && yDensity > 0)
                    {
                        dpiX = xDensity;
                        dpiY = yDensity;
                    }
                    else if (units == 2 && xDensity > 0 && yDensity > 0)
                    {
                        dpiX = xDensity * 2.54;
                        dpiY = yDensity * 2.54;
                    }
                }
                // APP14 Adobe (флаг инверсии для CMYK JPEG)
                else if (marker == 0xEE && segmentLength >= 14 &&
                         bytes[pos + 2] == 'A' && bytes[pos + 3] == 'd' &&
                         bytes[pos + 4] == 'o' && bytes[pos + 5] == 'b' && bytes[pos + 6] == 'e')
                {
                    hasAdobeApp14 = true;
                }
                // SOF0..SOF3
                else if (marker >= 0xC0 && marker <= 0xC3 && segmentLength >= 8)
                {
                    int bitsPerComponent = bytes[pos + 2];
                    int height = (bytes[pos + 3] << 8) | bytes[pos + 4];
                    int width = (bytes[pos + 5] << 8) | bytes[pos + 6];
                    int components = bytes[pos + 7];

                    string colorSpace = components switch
                    {
                        1 => "/DeviceGray",
                        4 => "/DeviceCMYK",
                        _ => "/DeviceRGB"
                    };

                    return new MinimalPdfImage(
                        width, height, dpiX, dpiY,
                        isJpeg: true,
                        bitsPerComponent: bitsPerComponent,
                        colorSpace: colorSpace,
                        compressedData: bytes,
                        invertCmyk: components == 4 && hasAdobeApp14);
                }

                pos += segmentLength;
            }

            throw new InvalidDataException("Unable to read JPEG header (SOF marker not found).");
        }

        private static MinimalPdfImage ParsePng(byte[] bytes, Func<byte[], byte[]> zlibCompressor)
        {
            int pos = 8;
            int width = 0, height = 0;
            byte bitDepth = 0, colorType = 0, interlace = 0;
            double dpiX = 96.0, dpiY = 96.0;
            byte[] palette = null;
            byte[] transparency = null;
            using var idatStream = new MemoryStream();

            while (pos + 12 <= bytes.Length)
            {
                uint length = ReadUInt32BE(bytes, pos);
                string type = $"{(char)bytes[pos + 4]}{(char)bytes[pos + 5]}{(char)bytes[pos + 6]}{(char)bytes[pos + 7]}";
                int dataStart = pos + 8;
                if (dataStart + length + 4 > bytes.Length)
                    break;

                if (type == "IHDR" && length >= 13)
                {
                    width = (int)ReadUInt32BE(bytes, dataStart);
                    height = (int)ReadUInt32BE(bytes, dataStart + 4);
                    bitDepth = bytes[dataStart + 8];
                    colorType = bytes[dataStart + 9];
                    interlace = bytes[dataStart + 12];
                }
                else if (type == "pHYs" && length >= 9)
                {
                    uint ppuX = ReadUInt32BE(bytes, dataStart);
                    uint ppuY = ReadUInt32BE(bytes, dataStart + 4);
                    byte unit = bytes[dataStart + 8];
                    if (unit == 1 && ppuX > 0 && ppuY > 0)
                    {
                        dpiX = ppuX * 0.0254;
                        dpiY = ppuY * 0.0254;
                    }
                }
                else if (type == "PLTE")
                {
                    palette = new byte[length];
                    Buffer.BlockCopy(bytes, dataStart, palette, 0, (int)length);
                }
                else if (type == "tRNS")
                {
                    transparency = new byte[length];
                    Buffer.BlockCopy(bytes, dataStart, transparency, 0, (int)length);
                }
                else if (type == "IDAT")
                {
                    idatStream.Write(bytes, dataStart, (int)length);
                }
                else if (type == "IEND")
                {
                    break;
                }

                pos = dataStart + (int)length + 4;
            }

            if (width <= 0 || height <= 0)
                throw new InvalidDataException("Invalid PNG IHDR chunk.");
            if (interlace != 0)
                throw new NotSupportedException("Interlaced (Adam7) PNG images are not supported.");

            byte[] compressedIdat = idatStream.ToArray();
            byte[] rawScanlines;
            using (var ms = new MemoryStream(compressedIdat, 2, compressedIdat.Length - 6))
            using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
            using (var outMs = new MemoryStream())
            {
                ds.CopyTo(outMs);
                rawScanlines = outMs.ToArray();
            }

            int channels = colorType switch
            {
                0 => 1, // Grayscale
                2 => 3, // RGB
                3 => 1, // Indexed (Palette)
                4 => 2, // Grayscale + Alpha
                6 => 4, // RGBA
                _ => throw new NotSupportedException($"Unsupported PNG color type: {colorType}")
            };

            if (bitDepth != 8)
                throw new NotSupportedException($"Only 8-bit PNG images are supported (got {bitDepth}-bit).");

            int stride = width * channels;
            byte[] unfiltered = UnfilterPngScanlines(rawScanlines, height, channels, stride);

            bool isGray = colorType == 0 || colorType == 4;
            bool hasAlpha = colorType == 4 || colorType == 6 || (colorType == 3 && transparency != null);
            int outColorChannels = isGray ? 1 : 3;

            byte[] colorData = new byte[width * height * outColorChannels];
            byte[] alphaData = hasAlpha ? new byte[width * height] : null;

            int srcIdx = 0;
            int dstColorIdx = 0;
            int dstAlphaIdx = 0;

            for (int i = 0; i < width * height; i++)
            {
                if (colorType == 0)
                {
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                }
                else if (colorType == 2)
                {
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                }
                else if (colorType == 3)
                {
                    int palIdx = unfiltered[srcIdx++];
                    colorData[dstColorIdx++] = palette[palIdx * 3];
                    colorData[dstColorIdx++] = palette[palIdx * 3 + 1];
                    colorData[dstColorIdx++] = palette[palIdx * 3 + 2];
                    if (hasAlpha)
                        alphaData[dstAlphaIdx++] = palIdx < transparency.Length ? transparency[palIdx] : (byte)255;
                }
                else if (colorType == 4)
                {
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    alphaData[dstAlphaIdx++] = unfiltered[srcIdx++];
                }
                else if (colorType == 6)
                {
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    colorData[dstColorIdx++] = unfiltered[srcIdx++];
                    alphaData[dstAlphaIdx++] = unfiltered[srcIdx++];
                }
            }
            byte[] compressedColor = zlibCompressor(colorData);
            byte[] compressedAlpha = hasAlpha ? zlibCompressor(alphaData) : null;
            return new MinimalPdfImage(
                width,
                height,
                dpiX,
                dpiY,
                isJpeg: false,
                bitsPerComponent: 8,
                colorSpace: isGray ? "/DeviceGray" : "/DeviceRGB",
                compressedData: compressedColor,
                alphaMaskCompressedData: compressedAlpha);
        }

        private static byte[] UnfilterPngScanlines(byte[] raw, int height, int bpp, int stride)
        {
            byte[] result = new byte[height * stride];
            int srcPos = 0;
            for (int y = 0; y < height; y++)
            {
                byte filterType = raw[srcPos++];
                int rowStart = y * stride;
                int prevRowStart = (y - 1) * stride;
                for (int x = 0; x < stride; x++)
                {
                    byte rawByte = raw[srcPos++];
                    byte a = x >= bpp ? result[rowStart + x - bpp] : (byte)0;
                    byte b = y > 0 ? result[prevRowStart + x] : (byte)0;
                    byte c = (x >= bpp && y > 0) ? result[prevRowStart + x - bpp] : (byte)0;
                    byte recon = filterType switch
                    {
                        0 => rawByte,
                        1 => (byte)(rawByte + a),
                        2 => (byte)(rawByte + b),
                        3 => (byte)(rawByte + ((a + b) >> 1)),
                        4 => (byte)(rawByte + PaethPredictor(a, b, c)),
                        _ => throw new InvalidDataException($"Invalid PNG filter type: {filterType}")
                    };
                    result[rowStart + x] = recon;
                }
            }
            return result;
        }

        private static byte PaethPredictor(byte a, byte b, byte c)
        {
            int p = a + b - c;
            int pa = Math.Abs(p - a);
            int pb = Math.Abs(p - b);
            int pc = Math.Abs(p - c);
            if (pa <= pb && pa <= pc) return a;
            if (pb <= pc) return b;
            return c;
        }

        private static uint ReadUInt32BE(byte[] b, int o)
            => ((uint)b[o] << 24) | ((uint)b[o + 1] << 16) | ((uint)b[o + 2] << 8) | b[o + 3];
    }
}