using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace SharpLayout.Tests
{
    public static class ImageUtil
    {
        public static void ImageDiff(string path1, byte[] image2, string diffPath)
        {
            using (var bitmap1 = new Bitmap(path1))
            using (var stream = new MemoryStream(image2))
            using (var bitmap2 = new Bitmap(stream))
            {
                var width = Math.Max(bitmap1.Width, bitmap2.Width);
                var height = Math.Max(bitmap1.Height, bitmap2.Height);
                using (var bitmap = new Bitmap(width, height))
                {
                    using (var imageData1 = bitmap1.LockBitsDisposable(new Rectangle(Point.Empty, bitmap1.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb))
                    using (var imageData2 = bitmap2.LockBitsDisposable(new Rectangle(Point.Empty, bitmap2.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb))
                    using (var imageData = bitmap.LockBitsDisposable(new Rectangle(Point.Empty, bitmap.Size), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb))
                    {
                        var values1 = new byte[imageData1.Stride * bitmap1.Height];
                        Marshal.Copy(imageData1.Scan0, values1, 0, values1.Length);
                        var values2 = new byte[imageData2.Stride * bitmap2.Height];
                        Marshal.Copy(imageData2.Scan0, values2, 0, values2.Length);
                        var values = new byte[Math.Max(values1.Length, values2.Length)];
                        for (var i = 0; i < values.Length; i += 4)
                        {
                            void SetRed(bool b)
                            {
                                if (b)
                                {
                                    values[i] = 75;
                                    values[i + 1] = 51;
                                    values[i + 2] = 244;
                                    values[i + 3] = 255;
                                }
                                else
                                {
                                    values[i] = 80;
                                    values[i + 1] = 192;
                                    values[i + 2] = 47;
                                    values[i + 3] = 255;
                                }
                            }
                            if (i < values1.Length && i < values2.Length)
                            {
                                if (values1[i] == values2[i] &&
                                    values1[i + 1] == values2[i + 1] &&
                                    values1[i + 2] == values2[i + 2])
                                {
                                    values[i] = values1[i];
                                    values[i + 1] = values1[i + 1];
                                    values[i + 2] = values1[i + 2];
                                    values[i + 3] = 51;
                                }
                                else
                                {
                                    SetRed(values2[i] == 255 &&
                                        values2[i + 1] == 255 &&
                                        values2[i + 2] == 255);
                                }
                            }
                            else
                            {
                                SetRed(i >= values1.Length);
                            }
                        }
                        Marshal.Copy(values, 0, imageData.Scan0, values.Length);
                    }
                    using (var bitmap3 = new Bitmap(width, height))
                    {
                        using (var graphics = System.Drawing.Graphics.FromImage(bitmap3))
                        {
                            graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, width, height);
                            graphics.DrawImage(bitmap, new Point(0, 0));
                        }
                        bitmap3.Save(diffPath);
                    }
                }
            }
        }

        public static DisposableImageData LockBitsDisposable(this Bitmap bitmap, Rectangle rect, ImageLockMode flags, PixelFormat format)
        {
            return new DisposableImageData(bitmap, rect, flags, format);
        }
    }

    public class DisposableImageData : IDisposable
    {
        private readonly Bitmap bitmap;
        private readonly BitmapData bitmapData;

        internal DisposableImageData(Bitmap bitmap, Rectangle rect, ImageLockMode flags, PixelFormat format)
        {
            this.bitmap = bitmap;
            bitmapData = bitmap.LockBits(rect, flags, format);
        }

        public void Dispose() => bitmap.UnlockBits(bitmapData);
        public IntPtr Scan0 => bitmapData.Scan0;
        public int Stride => bitmapData.Stride;
        public int Width => bitmapData.Width;
        public int Height => bitmapData.Height;
        public PixelFormat PixelFormat => bitmapData.PixelFormat;
        public int Reserved => bitmapData.Reserved;
    }
}