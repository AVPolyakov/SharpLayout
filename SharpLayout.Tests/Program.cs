using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.MetaData;
using SixLabors.ImageSharp.PixelFormats;
using IImage = PdfSharp.Drawing.IImage;

namespace SharpLayout.Tests
{
    public class ImageFactory : IImageFactory
    {
        public IImage FromFile(string path)
        {            
            return new PngImage(SixLabors.ImageSharp.Image.Load(path), 
                SixLabors.ImageSharp.Image.DetectFormat(path));
        }

        public IImage FromStream(Stream stream)
        {
            return new PngImage(SixLabors.ImageSharp.Image.Load(stream),
                SixLabors.ImageSharp.Image.DetectFormat(stream));
        }

        public class PngImage : IImage
        {
            private readonly Image<Rgba32> _image;
            private readonly IImageFormat _imageFormat;

            public PngImage(Image<Rgba32> image, IImageFormat imageFormat)
            {
                _image = image;
                _imageFormat = imageFormat;
            }

            public void Dispose() => _image.Dispose();
            public XImageFormat ImageFormat
            {
                get
                {
                    if (_imageFormat == PngFormat.Instance)
                        return XImageFormat.Png;
                    if (_imageFormat == JpegFormat.Instance)
                        return XImageFormat.Jpeg;
                    throw new ArgumentException(nameof(_imageFormat));
                }
            }

            public int Width => _image.Width;
            public int Height => _image.Height;

            public double HorizontalResolution => GetResolution(_image.MetaData.HorizontalResolution);

            public double VerticalResolution => GetResolution(_image.MetaData.VerticalResolution);

            public void AddPdfElements(PdfDictionary.DictionaryElements elements)
            {
                elements[PdfImage.Keys.ColorSpace] = new PdfName("/DeviceRGB");
            }

            public BitmapReader GetBitmapReader()
            {
                return new BitmapReader.TrueColor(3, 8, false);
            }

            public void SaveToBmp(MemoryStream memory)
            {
                _image.SaveAsBmp(memory);
            }

            private double GetResolution(double resolution)
            {
                switch (_image.MetaData.ResolutionUnits)
                {
                    case PixelResolutionUnit.PixelsPerInch:
                        return resolution;
                    case PixelResolutionUnit.PixelsPerCentimeter:
                        return resolution * 2.54;
                    case PixelResolutionUnit.PixelsPerMeter:
                        return resolution * 2.54 / 100;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(_image.MetaData.ResolutionUnits));
                }
            }
        }
    }

    public class FontResolver : IFontResolver
    {
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            switch (familyName)
            {
                case "Times New Roman":
                    if (isBold)
                    {
                        if (isItalic)
                        {
                        }
                        else
                            return new FontResolverInfo(TimesNewRomanBold);
                    }
                    else
                    {
                        if (isItalic)
                        {
                        }
                        else
                            return new FontResolverInfo(TimesNewRoman);
                    }
                    break;
            }
            throw new Exception("Font not found");
        }

        public byte[] GetFont(string faceName)
        {
            switch (faceName)
            {
                case TimesNewRoman:
                    return File.ReadAllBytes(@"E:\Temp\Fonts\times.ttf");
                case TimesNewRomanBold:
                    return File.ReadAllBytes(@"E:\Temp\Fonts\timesbd.ttf");
            }
            throw new InvalidOperationException();
        }

        private const string TimesNewRoman = "Times New Roman";
        private const string TimesNewRomanBold = "Times New Roman,Bold";
    }

    static class Program
    {
        static void Main()
        {
            Document.CollectCallerInfo = true;
            GlobalFontSettings.FontResolver = new FontResolver();
            ImageSettings.ImageFactory = new ImageFactory();

            try
            {
                var document = new Document {
                    //CellsAreHighlighted = true,
                    //R1C1AreVisible = true,
                    //ParagraphsAreHighlighted = true,
                    //CellLineNumbersAreVisible = true,
                    //ExpressionVisible = true,
                };
                PaymentOrder.AddSection(document, new PaymentData {IncomingDate = DateTime.Now, OutcomingDate = DateTime.Now});
                //Svo.AddSection(document);
                //ContractDealPassport.AddSection(document);
                //LoanAgreementDealPassport.AddSection(document);

                //document.SavePng(0, "Temp.png", 120).StartLiveViewer(true);

                //Process.Start(document.SavePng(0, "Temp2.png")); //open with Paint.NET
                Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            }
            catch (Exception e)
            {
                ShowException(e);
            }
        }

        private static void ShowException(Exception e)
        {
            var document = new Document();
            var settings = new PageSettings();
            settings.LeftMargin = settings.TopMargin = settings.RightMargin = settings.BottomMargin = Util.Cm(0.5);
            document.Add(new Section(settings).Add(new Paragraph()
                .Add($"{e}", new Font("Consolas", 9.5, XFontStyle.Regular, Styles.PdfOptions))));
            //document.SavePng(0, "Temp.png", 120).StartLiveViewer(true, false);
            Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
        }

        public static void StartLiveViewer(this string fileName, bool alwaysShowWindow, bool findId = true)
        {
            var processes = Process.GetProcessesByName("LiveViewer");
            if (processes.Length <= 0)
            {
                string arguments;
                if (findId && Ide == vs)
                {
                    const string solutionName = "SharpLayout";
                    var firstOrDefault = Process.GetProcesses().FirstOrDefault(p => p.ProcessName == "devenv" &&
                        p.MainWindowTitle.Contains(solutionName));
                    if (firstOrDefault != null)
                        arguments = $" {firstOrDefault.Id}";
                    else
                        arguments = "";
                }
                else
                    arguments = "";
                Process.Start("LiveViewer", fileName + " " + Ide + arguments);
            }
            else
            {
                if (alwaysShowWindow)
                    SetForegroundWindow(processes[0].MainWindowHandle);
            }
        }

        private static string Ide => vs;

        private const string vs = "vs";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
