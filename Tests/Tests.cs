﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Examples;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using QRCoder;
using Resources;
using SharpLayout;
using SharpLayout.ImageRendering;
using Tests.Images;
using Xunit;
using static SharpLayout.Direction;
using static SharpLayout.InlineVerticalAlign;
using static SharpLayout.Util;
using static Examples.Styles;
using static Resources.FontFamilies;
using static SharpLayout.Dpi254;

namespace Tests
{
    public class Tests
    {
        static Tests()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }
        
        [Fact]
        public void QRCode()
        {
            var qrGenerator = new QRCodeGenerator();
            var qrCodeData = qrGenerator.CreateQrCode("https://github.com/AVPolyakov/SharpLayout", QRCodeGenerator.ECCLevel.Q);
            var qrCode = new PngByteQRCode(qrCodeData);
            var bytes = qrCode.GetGraphic(20, drawQuietZones: false);

            var size = Cm(2);
            
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(size);
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Content(() => new MemoryStream(bytes))
                    .Width(size).Height(size));
            }

            //StartProcess(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            Assert(nameof(QRCode), document.CreatePng().Item1);
        }

        // ReSharper disable once UnusedMember.Local
        private static void StartProcess(string fileName)
        {
            Process.Start(new ProcessStartInfo("cmd", $"/c start {fileName}") {CreateNoWindow = true});
        }
        
        [Fact]
        public void EmptyFirstPageHeader()
        {
            var document = new Document();
            var pageSettings = new PageSettings {
                TopMargin = 0,
                BottomMargin = Cm(1),
                LeftMargin = Cm(1),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            section.SetEmptyFirstPageHeaders();
            {
                var table = section.AddHeader().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Cm(0.5));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(c => $"Page {c.PageNumber} of {c.PageCount}"));
            }
            section.AddPageBreak();
            section.AddPageBreak();
            Assert(nameof(EmptyFirstPageHeader), document.CreatePng().Item1);
        }

        [Fact]
        public void DifferentFooters()
        {
            var document = new Document();
            var pageSettings = new PageSettings {
                TopMargin = 0,
                BottomMargin = 0,
                LeftMargin = Cm(1),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            {
                var table = section.AddFirstPageFooter().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Px(20));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"First footer
First footer
First footer
")
                    .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
            }
            {
                var table = section.AddEvenPageFooter().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Px(20));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Second footer
")
                    .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
            }
            {
                var table = section.AddFooter().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Px(20));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Other footer
Other footer
")
                    .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
            }
            section.Add(new Paragraph().TextIndent(Cm(1)).Alignment(HorizontalAlign.Justify)
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ", Styles.TimesNewRoman10));
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 1, row {i}"));
                }
            }
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 2, row {i}"));
                }
            }
            Assert(nameof(DifferentFooters), document.CreatePng().Item1);
        }
        
        [Fact]
        public void DifferentPageHeaders()
        {
            var document = new Document();
            var pageSettings = new PageSettings {
                TopMargin = 0,
                BottomMargin = Cm(1),
                LeftMargin = Cm(1),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            {
                var table = section.AddFirstPageHeader().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Cm(0.5));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"First header
First header
First header
")
                    .Add(c => $"Page {c.PageNumber} of {c.PageCount}"));
            }
            {
                var table = section.AddEvenPageHeader().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Cm(0.5));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Second header
")
                    .Add(c => $"Page {c.PageNumber} of {c.PageCount}"));
            }
            {
                var table = section.AddHeader().Font(Styles.TimesNewRoman10)
                    .Margin(Top | Bottom, Cm(0.5));
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Other header
Other header
")
                    .Add(c => $"Page {c.PageNumber} of {c.PageCount}"));
            }
            section.Add(new Paragraph().TextIndent(Cm(1)).Alignment(HorizontalAlign.Justify)
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    Styles.TimesNewRoman10));
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 1, row {i}"));
                }
            }
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 2, row {i}"));
                }
            }
            Assert(nameof(DifferentPageHeaders), document.CreatePng().Item1);
        }
        
        [Fact]
        public void DifferentPageHeaders2()
        {
            var document = new Document();
            var pageSettings = new PageSettings {
                TopMargin = 0,
                BottomMargin = 0,
                LeftMargin = Cm(1),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            {
                var table = section.AddFirstPageHeader().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"First header"));
            }
            {
                var table = section.AddEvenPageHeader().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Second header
Second header"));
            }
            {
                var table = section.AddHeader().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(@"Other header
Other header
Other header"));
            }
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10).KeepWithNext(true);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 1, row {i}"));
                }
            }
            {
                var table = section.AddTable().Font(Styles.TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 160; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 2, row {i}"));
                }
            }
            Assert(nameof(DifferentPageHeaders2), document.CreatePng().Item1);
        }

        [Fact]
        public void SectionGroup()
        {
            var document = new Document();
            var sectionGroup = document.Add(new SectionGroup());
            {
                var section = sectionGroup.Add(new Section(new PageSettings()));
                section.Add(new Paragraph().Add("Test", Styles.TimesNewRoman10));
                var footers = section.AddFooter().Font(Styles.TimesNewRoman10);
                var c1 = footers.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = footers.AddRow().Height(Px(200));
                r1[c1].Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add(c => $"{c.PageNumber} из {c.PageCount}"));
            }
            {
                var section = sectionGroup.Add(new Section(new PageSettings()));
                section.Add(new Paragraph().Add("Test2", Styles.TimesNewRoman10));
                var footers = section.AddFooter().Font(Styles.TimesNewRoman10);
                var c1 = footers.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = footers.AddRow().Height(Px(200));
                r1[c1].Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add(c => $"{c.PageNumber} of {c.PageCount}"));
                section.AddPageBreak();
            }
            Assert(nameof(SectionGroup), document.CreatePng().Item1);
        }

        [Fact]
        public void KeepLinesTogether()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            for (var i = 0; i < 15; i++)
                section.Add(new Paragraph().KeepLinesTogether(true)
                    .Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))                    
                    .Add("Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated. " +
                         "Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated. " +
                         "Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated.",
                        Styles.TimesNewRoman10));
            Assert(nameof(KeepLinesTogether), document.CreatePng().Item1);
        }

        [Fact]
        public void AddPageBreak()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var footers = section.AddFooter().Font(Styles.TimesNewRoman10);
                var c1 = footers.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = footers.AddRow().Height(Px(200));
                r1[c1].Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add(c => $"{c.PageNumber} из {c.PageCount}"));
            }
            section.Add(new Paragraph().Add("Test1", Styles.TimesNewRoman10));
            section.AddPageBreak();
            section.Add(new Paragraph().Add("Test2", Styles.TimesNewRoman10));
            section.AddPageBreak();
            section.AddPageBreak();
            section.Add(new Paragraph().Add("Test3", Styles.TimesNewRoman10));
            Assert(nameof(AddPageBreak), document.CreatePng().Item1);
        }

        [Fact]
        public void SecuritySettings()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            section.Add(new Paragraph().Add("Test", Styles.TimesNewRoman10));

            //byte[] bytes;
            using (var pdfDocument = new PdfDocument())
            {
				//http://www.pdfsharp.com/PDFsharp/index.php%3Foption%3Dcom_content%26task%3Dview%26id%3D36%26Itemid%3D47
	            var securitySettings = pdfDocument.SecuritySettings;
 
	            // Setting one of the passwords automatically sets the security level to
	            // // PdfDocumentSecurityLevel.Encrypted128Bit.
	            //securitySettings.UserPassword  = "user";
	            //securitySettings.OwnerPassword = "owner";
 
	            // DonÂ´t use 40 bit encryption unless needed for compatibility reasons
	            // //securitySettings.DocumentSecurityLevel = PdfDocumentSecurityLevel.Encrypted40Bit;
 
	            // Restrict some rights.
	            securitySettings.PermitAccessibilityExtractContent = false;
	            securitySettings.PermitAnnotations = false;
	            securitySettings.PermitAssembleDocument = false;
	            securitySettings.PermitExtractContent = false;
	            securitySettings.PermitFormsFill = false;
	            securitySettings.PermitFullQualityPrint = false;
	            securitySettings.PermitModifyDocument = false;
	            //securitySettings.PermitPrint = false;

	            document.Render(pdfDocument);
	            using (var stream = new MemoryStream())
	            {
		            pdfDocument.Save(stream);
		            //bytes = stream.ToArray();
	            }
            }

            //var path = $"Temp_{Guid.NewGuid():N}.pdf";
			//File.WriteAllBytes(path, bytes);
			//Process.Start(path);
        }

        [Fact]
        public void Image()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var table = section.AddTable().Border(BorderWidth);
            var c1 = table.AddColumn(Px(800));
            var r1 = table.AddRow();
            r1[c1].Add(new Image()
	            .Content(Image1));
            var r2 = table.AddRow();
            r2[c1].Add(new Image()
	            .Content(Image2));
            Assert(nameof(Image), document.CreatePng().Item1);
        }

        private static FileStream Image1() => File.OpenRead(@"Images\Image1.png");
        private static FileStream Image2() => File.OpenRead(@"Images\Image2.jpg");

        [Fact]
        public void Image_Alignments()
        {
	        var document = new Document();
	        var section = document.Add(new Section(new PageSettings()));
	        var table = section.AddTable().Border(BorderWidth);
	        var c1 = table.AddColumn(Px(500));
	        var c2 = table.AddColumn(Px(500));
	        var c3 = table.AddColumn(Px(500));
	        var r1 = table.AddRow();
	        r1[c1].Add(new Image()
		        .Content(Image1));
	        var r2 = table.AddRow();
	        r2[c2].Add(new Image().Alignment(HorizontalAlign.Center)
		        .Content(Image1));
	        var r3 = table.AddRow();
	        r3[c3].Add(new Image().Alignment(HorizontalAlign.Right)
		        .Content(ResourceImage1));
            var r4 = table.AddRow();
            r4[c3].Add(new Image().Height(Cm(1)).Width(Cm(3)).Alignment(HorizontalAlign.Right)
                .Content(ResourceImage1));
            Assert(nameof(Image_Alignments), document.CreatePng().Item1);
        }

        private static Stream ResourceImage1()
        {
            var anchorType = typeof(ResourceImage1);
            return anchorType.Assembly.GetManifestResourceStream(anchorType.FullName + ".png");
        }

        [Fact]
        public void VectorImage()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var info = blueRabbitInfo.Value;
            {
                var table = section.AddTable();
                var scale = 0.7;
                var c1 = table.AddColumn(info.PointWidth * scale);
                table.AddColumn(Px(500));
                table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(BlueRabbit));
            }
            {
                var table = section.AddTable().Margin(Top, Cm(0.5));
                var scale = 0.3;
                var c1 = table.AddColumn();
                var c2 = table.AddColumn();
                var c3 = table.AddColumn();
                table.Columns.ToArray().Distribute(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(BlueRabbit));
                var r2 = table.AddRow();
                r2[c2].Add(new Image().Alignment(HorizontalAlign.Center)
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(BlueRabbit));
                var r3 = table.AddRow();
                r3[c3].Add(new Image().Alignment(HorizontalAlign.Right)
                    .Width(info.PointWidth * scale).Height(info.PointHeight * scale)
                    .Content(BlueRabbit));
            }
            //Process.Start(document.SavePdf($"Temp_{Guid.NewGuid():N}.pdf"));
            Assert(nameof(VectorImage), document.CreatePng().Item1);
        }

        private static readonly Lazy<ImageInfo> blueRabbitInfo = new Lazy<ImageInfo>(() => {
            var bytes = File.ReadAllBytes(@"Images\blue-rabbit.pdf");
            using (var stream = new MemoryStream(bytes))
            using (var xImage = XImage.FromStream(stream))
                return new ImageInfo(bytes, xImage.PointWidth, xImage.PointHeight);
        });

        private static MemoryStream BlueRabbit() => new MemoryStream(blueRabbitInfo.Value.Bytes);

        [Fact]
        public void Footnotes()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(0.7)
            }));
            var footnoteFont = new Font(TimesNewRoman, 10, XFontStyle.Regular, PdfOptions);
            section.AddFootnoteSeparator(FootnoteSeparator());
            {
                var table = section.AddTable()
                    .Font(new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions));
                var c1 = table.AddColumn(Px(500));
                for (var i = 0; i < 3; i++)
                    table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                        .Add("Text"));
                table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                    .Add("First text")
                    .Add(new Span("1").InlineVerticalAlign(Super)
                        .AddFootnote(new Paragraph()
                                .Add(new Span("1", footnoteFont).InlineVerticalAlign(Super))
                                .Add(" Footnote1", footnoteFont))));
                for (var i = 0; i < 3; i++)
                    table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                        .Add("Text"));
                table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                    .Add("Second text")
                    .Add(new Span("2").InlineVerticalAlign(Super)
                        .AddFootnote(new Paragraph()
                                .Add(new Span("2", footnoteFont).InlineVerticalAlign(Super))
                                .Add(" Footnote2", footnoteFont))));
                for (var i = 0; i < 65; i++)
                    table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                        .Add("Text"));
                table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                    .Add("Third text")
                    .Add(new Span("3").InlineVerticalAlign(Super)
                        .AddFootnote(new Paragraph()
                                .Add(new Span("3", footnoteFont).InlineVerticalAlign(Super))
                                .Add(" Footnote3", footnoteFont))));
                for (var i = 0; i < 100; i++)
                    table.AddRow()[c1].Add(new Paragraph().Margin(Left | Right, Px(5))
                        .Add("Text"));
            }
            Assert(nameof(Footnotes), document.CreatePng().Item1);
        }

        [Fact]
        public void Footnotes2()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(0.7)
            }));
            var footnoteFont = new Font(TimesNewRoman, 10, XFontStyle.Regular, PdfOptions);
            var font = new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions);
            section.AddFootnoteSeparator(FootnoteSeparator());
            for (var i = 0; i < 54; i++)
                section.Add(new Paragraph()
                    .Add("Text", font));
            var table = section.AddTable().Font(font).Border(BorderWidth);
            var c1 = table.AddColumn(Cm(5));
            table.AddRow().KeepWith(2)[c1].Add(new Paragraph()
                .Add("r1c1")
                .Add(new Span("1").InlineVerticalAlign(Super)
                    .AddFootnote(new Paragraph()
                        .Add(new Span("1", footnoteFont).InlineVerticalAlign(Super))
                        .Add("FootnoteFont1", footnoteFont))));
            table.AddRow()[c1].Add(new Paragraph()
                .Add("r2c1")
                .Add(new Span("2").InlineVerticalAlign(Super)
                    .AddFootnote(new Paragraph()
                        .Add(new Span("2", footnoteFont).InlineVerticalAlign(Super))
                        .Add("FootnoteFont2", footnoteFont))));
            Assert(nameof(Footnotes2), document.CreatePng().Item1);
        }

        private static Table FootnoteSeparator()
        {
            var table = new Table().Margin(Top | Bottom, Px(15));
            var c = table.AddColumn(Cm(5));
            var r = table.AddRow();
            r[c].Border(Top);
            return table;
        }

        [Fact]
	    public void UnderlineText()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings()));
		    var table = section.AddTable();
		    var c1 = table.AddColumn(Px(200));
		    var r1 = table.AddRow();
		    r1[c1].Add(new Paragraph().Add("Test", new Font(TimesNewRoman, 10, XFontStyle.Underline, PdfOptions)));
		    Assert(nameof(UnderlineText), document.CreatePng().Item1);
		}

	    [Fact]
	    public void AddParagraph_TrailingSpace()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings()));
		    var table = section.AddTable();
		    var c1 = table.AddColumn(Px(200));
		    var r1 = table.AddRow();
		    r1[c1].Add(new Paragraph().Add("wwxxxxxxxx ", Styles.TimesNewRoman10));
		    Assert(nameof(AddParagraph_TrailingSpace), document.CreatePng().Item1);
		}

	    [Fact]
	    public void AddParagraph_TrailingSpace2()
	    {
			var document = new Document();
			var section = document.Add(new Section(new PageSettings()));
		    var table = section.AddTable();
		    var c1 = table.AddColumn(Px(200));
		    var r1 = table.AddRow();
		    r1[c1].Add(new Paragraph().Add(@"line1
wwxxxxxxxx 
line3", Styles.TimesNewRoman10));
		    Assert(nameof(AddParagraph_TrailingSpace2), document.CreatePng().Item1);
		}

	    [Fact]
	    public void AddParagraph_LineBreak()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings()));
		    section.Add(new Paragraph().Add(@"qwe

qwe2
qwe3
", Styles.TimesNewRoman10))
			    .Add(new Paragraph().Add(@"test", Styles.TimesNewRoman10));
		    Assert(nameof(AddParagraph_LineBreak), document.CreatePng().Item1);
		}

	    [Fact]
	    public void AddParagraph_EmptyString()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings()));
		    section.Add(new Paragraph().Add("", Styles.TimesNewRoman10))
			    .Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
		    Assert(nameof(AddParagraph_EmptyString), document.CreatePng().Item1);
		}

	    [Fact]
	    public void ParagraphWithSpace()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings()));
		    section.Add(new Paragraph().Add(" ", Styles.TimesNewRoman10));
		    document.CreatePng();
		}

        [Fact]
        public void UnderlinedSpace()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var timesNewRoma = new Font(TimesNewRoman, 9, XFontStyle.Underline, PdfOptions);
            section.Add(new Paragraph().Add("    one     ", timesNewRoma).Add("\u200B", timesNewRoma));
            Assert(nameof(UnderlinedSpace), document.CreatePng().Item1);
        }

        [Fact]
        public void Checkbox()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var timesNewRoma = new Font(TimesNewRoman, 9, XFontStyle.Regular, PdfOptions);
            var wingdings = new Font(Wingdings, 11, XFontStyle.Bold, PdfOptions);
            var wingdings2 = new Font(Wingdings2, 11, XFontStyle.Bold, PdfOptions);
            section.Add(new Paragraph().Add("\u00A3", wingdings2)
                    .Add(@" Одна тысяча
одна тысяча", timesNewRoma))
                .Add(new Paragraph().Add("\u0053", wingdings2)
                    .Add(@" Две тысячи
две тысячи", timesNewRoma))
                .Add(new Paragraph().Margin(Top, Cm(2)).Add("\u00FE\u00FD\u00A8", wingdings)
                    .Add("\u0054\u0053\u00A3\u0052\u0051", wingdings2)
                    .Add("Одна тысяча", timesNewRoma));
            Assert(nameof(Checkbox), document.CreatePng().Item1);
        }

        [Fact]
        public void Superscript()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var table = section.AddTable().Font(new Font(TimesNewRoman, 20, XFontStyle.Regular, PdfOptions));
                var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Margin(Bottom, Px(10))
                    .Add("You can set the ")
                    .Add(new Span("subscript").InlineVerticalAlign(Sub))
                    .Add(" or ")
                    .Add(new Span("superscript").InlineVerticalAlign(Super))
                    .Add("."));
            }
            {
                var table = section.AddTable()
                    .Border(BorderWidth)
                    .Font(new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions));
                table.AddColumn(Px(300));
                var c2 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c2].Add(new Paragraph().Margin(Left | Right, Px(5))
                    .Add("Some text").Add(new Span("1").InlineVerticalAlign(Super)));
            }
            {
                var table = section.AddTable().Font(new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions));
                var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add(new Span("1").InlineVerticalAlign(Super)).Add(" Footnote text"));
            }
            Assert(nameof(Superscript), document.CreatePng().Item1);
        }

        [Fact]
        public void Superscript_LongText()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            const string note1 = "1";
            const string note2 = "2";
            const string note3 = "3";
            {
                var font = new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions);
                section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).Margin(Bottom, Px(30))
                    .Add("Choose composition first when creating new classes from existing classes. " +
                        "Only if inheritance", font)
                    .Add(new Span(note1, font).InlineVerticalAlign(Super))
                    .Add(" is required by your design should it be used. If you use inheritance", font)
                    .Add(new Span(note2, font).InlineVerticalAlign(Super))
                    .Add(" where composition will work, your designs will become needlessly complicated", font)
                    .Add(new Span(note3, font).InlineVerticalAlign(Super))
                    .Add(".", font));
            }
            {
                var font = new Font(TimesNewRoman, 10, XFontStyle.Regular, PdfOptions);
                section.Add(new Paragraph()
                    .Add(new Span(note1, font).InlineVerticalAlign(Super))
                    .Add(" Text of first footnote", font));
                section.Add(new Paragraph()
                    .Add(new Span(note2, font).InlineVerticalAlign(Super))
                    .Add(" Text of second footnote", font));
                section.Add(new Paragraph()
                    .Add(new Span(note3, font).InlineVerticalAlign(Super))
                    .Add(" Text of third footnote", font));
            }
            Assert(nameof(Superscript_LongText), document.CreatePng().Item1);
        }

        [Fact]
        public void ParagraphByPages()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                BottomMargin = Cm(2)
            }));
            for (var i = 0; i < 20; i++)
            {
                var paragraph = new Paragraph()
                    .Alignment(HorizontalAlign.Justify)
                    .TextIndent(Cm(1))
                    .Add("Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated. " +
                         "Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated. " +
                         "Choose composition first when creating new classes from existing classes. Only if " +
                         "inheritance is required by your design should it be used. If you use inheritance where " +
                         "composition will work, your designs will become needlessly complicated.",
                        Styles.TimesNewRoman10);
                if (i == 10)
                {
                    var noteTable = new Table().Font(Styles.TimesNewRoman10);
                    var c = noteTable.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                    var r = noteTable.AddRow();
                    r[c].Add(new Paragraph()
                        .Add(new Span("1").InlineVerticalAlign(Super))
                        .Add(" Footnote1"));
                    paragraph.Add(new Span("1", Styles.TimesNewRoman10).InlineVerticalAlign(Super)
                        .AddFootnote(noteTable));
                }
                section.Add(paragraph);
            }
            Assert(nameof(ParagraphByPages), document.CreatePng().Item1);            
        }

	    [Fact]
	    public void RowHeightGreaterThanPageHeight()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings {
				TopMargin = 0
		    }));
		    section.Add(new Paragraph()
			    .Add("1 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
		    var table = section.AddTable().Border(BorderWidth);
		    var c1 = table.AddColumn(Px(500));
		    var r1 = table.AddRow();
		    r1.Height(Px(2700+500));
		    r1[c1].Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
		    Assert(nameof(RowHeightGreaterThanPageHeight), document.CreatePng().Item1);
	    }

	    [Fact]
	    public void KeepWithNext()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings {
				TopMargin = 0
		    }));
		    section.Add(new Paragraph()
			    .Add("1 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
		    section.Add(new Paragraph().KeepWithNext(true)
			    .Add("2 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
		    var table = section.AddTable().Border(BorderWidth);
		    var c1 = table.AddColumn(Px(500));
		    var r1 = table.AddRow();
		    r1.Height(Px(2700));
		    r1[c1].Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
		    Assert(nameof(KeepWithNext), document.CreatePng().Item1);
	    }

	    [Fact]
	    public void KeepWithNext_DifferentPageHeaders()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings {
				TopMargin = 0
		    }));
            {
                var table = section.AddFirstPageHeader();
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph().Add(@"First header
First header
First header", Styles.TimesNewRoman10));
            }
            {
                var table = section.AddEvenPageHeader();
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph().Add("Second header", Styles.TimesNewRoman10));
            }
            {
                var table = section.AddHeader();
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph().Add(@"Other header
Other header", Styles.TimesNewRoman10));
            }
            section.Add(new Paragraph()
			    .Add("1 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
		    section.Add(new Paragraph().KeepWithNext(true)
			    .Add("2 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
            {
                var table = section.AddTable().Border(BorderWidth);
                var c1 = table.AddColumn(Px(500));
                var r1 = table.AddRow();
                r1.Height(Px(2700-200));
                r1[c1].Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
            }
            section.Add(new Paragraph()
                .Add("3 Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    Styles.TimesNewRoman10));
            Assert(nameof(KeepWithNext_DifferentPageHeaders), document.CreatePng().Item1);
	    }
        
        [Fact]
	    public void KeepWith()
	    {
		    var document = new Document();
		    var section = document.Add(new Section(new PageSettings {
				TopMargin = 0,
				BottomMargin = 0
		    }));
		    section.Add(new Paragraph().KeepWithNext(true)
			    .Add("2 Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. " +
				    "Choose composition first when creating new classes from existing classes. Only if " +
				    "inheritance is required by your design should it be used. If you use inheritance where " +
				    "composition will work, your designs will become needlessly complicated. ",
				    Styles.TimesNewRoman10));
		    {
			    var table = section.AddTable().Border(BorderWidth);
			    table.AddColumn(Px(500));
			    table.AddRow().Height(Px(2600));
		    }
		    {
			    var table = section.AddTable().Border(BorderWidth);
			    table.AddColumn(Px(500));
			    table.AddRow().Height(Px(50)).KeepWith(4);
			    table.AddRow().Height(Px(50));
			    table.AddRow().Height(Px(50));
		    }
		    Assert(nameof(KeepWith), document.CreatePng().Item1);
	    }

        [Fact]
        public void TableBottomMargin()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var table = section.AddTable().Margin(Bottom, Px(200));
                var c1 = table.AddColumn(Px(200));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph().Add("Test1", Styles.TimesNewRoman10));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(200));
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph().Add("Test2", Styles.TimesNewRoman10));
            }
            Assert(nameof(TableBottomMargin), document.CreatePng().Item1);
        }

        [Fact]
        public void PageHeader()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var header = section.AddHeader();
                var c1 = header.AddColumn(Cm(5));
                var r1 = header.AddRow().Height(Px(300));
                r1[c1].Add(new Paragraph().Add("Заголовок", Styles.TimesNewRoman10));
            }
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(651));
                var c2 = table.AddColumn(Px(400));
                for (var i = 0; i < 200; i++)
                {
                    var r = table.AddRow();
                    if (i == 0)
                    {
                        r.TableHeader(true);
                        r[c1].Border(Top | Left | Right, BorderWidth * 1.5);
                        r[c2].Border(Top | Right | Bottom, BorderWidth * 1.5).Rowspan(2).VerticalAlign(VerticalAlign.Center)
                            .Add(new Paragraph().Alignment(HorizontalAlign.Center).Add("second", Styles.TimesNewRoman10));
                    }
                    else if (i == 1)
                    {
                        r.TableHeader(true);
                        r[c1].Border(Bottom | Left | Right, BorderWidth * 1.5);
                    }
                    else
                    {
                        r[c1].Border(Bottom | Left | Right, BorderWidth);
                        r[c2].Border(Bottom | Right, BorderWidth);
                    }
                    r[c1]
                        .Add(new Paragraph().Add($"first {i}", Styles.TimesNewRoman10));
                }
            }
            Assert(nameof(PageHeader), document.CreatePng().Item1);
        }

        [Fact]
        public void PageFooter()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            {
                var footers = section.AddFooter().Font(Styles.TimesNewRoman10);
                var c1 = footers.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = footers.AddRow().Height(Px(700));
                r1[c1].Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add(c => $"{c.PageNumber} из {c.PageCount}"));
            }
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(651));
                var c2 = table.AddColumn(Px(400));
                for (var i = 0; i < 200; i++)
                {
                    var r = table.AddRow();
                    if (i == 0)
                    {
                        r.TableHeader(true);
                        r[c1].Border(Top | Left | Right, BorderWidth * 1.5);
                        r[c2].Border(Top | Right | Bottom, BorderWidth * 1.5).Rowspan(2).VerticalAlign(VerticalAlign.Center)
                            .Add(new Paragraph().Alignment(HorizontalAlign.Center).Add("second", Styles.TimesNewRoman10));
                    }
                    else if (i == 1)
                    {
                        r.TableHeader(true);
                        r[c1].Border(Bottom | Left | Right, BorderWidth * 1.5);
                    }
                    else
                    {
                        r[c1].Border(Bottom | Left | Right, BorderWidth);
                        r[c2].Border(Bottom | Right, BorderWidth);
                    }
                    r[c1]
                        .Add(new Paragraph().Add($"first {i}", Styles.TimesNewRoman10));
                }
            }
            Assert(nameof(PageFooter), document.CreatePng().Item1);
        }

        [Fact]
        public void TableHeader()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            var table = section.AddTable();
            var c1 = table.AddColumn(Px(651));
            var c2 = table.AddColumn(Px(400));
            for (var i = 0; i < 200; i++)
            {
                var r = table.AddRow();
                if (i == 0)
                {
                    r.TableHeader(true);
                    r[c1].Border(Top | Left | Right, BorderWidth * 1.5);
                    r[c2].Border(Top | Right | Bottom, BorderWidth * 1.5).Rowspan(2).VerticalAlign(VerticalAlign.Center)
                        .Add(new Paragraph().Alignment(HorizontalAlign.Center).Add("second", Styles.TimesNewRoman10));
                }
                else if (i == 1)
                {
                    r.TableHeader(true);
                    r[c1].Border(Bottom | Left | Right, BorderWidth * 1.5);
                }
                else
                {
                    r[c1].Border(Bottom | Left | Right, BorderWidth);
                    r[c2].Border(Bottom | Right, BorderWidth);
                }
                r[c1]
                    .Add(new Paragraph().Add($"first {i}", Styles.TimesNewRoman10));
            }
            Assert(nameof(TableHeader), document.CreatePng().Item1);
        }

        [Fact]
        public void DashBorders()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var table = section.AddTable();
            var c1 = table.AddColumn(Px(651));
            var r1 = table.AddRow();
            var r2 = table.AddRow();
            var r3 = table.AddRow();
            r1[c1].Border(Bottom, new XPen(XColors.Black, BorderWidth * 2) {DashStyle = XDashStyle.Dash})
                .Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
            //How to: Draw a Custom Dashed Line https://docs.microsoft.com/en-us/dotnet/framework/winforms/advanced/how-to-draw-a-custom-dashed-line
            r2[c1].Border(Bottom, new XPen(XColors.Black, BorderWidth * 2) {DashPattern = new[] {3d, 3d}})
                .Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
            r3[c1].Border(Bottom, new XPen(XColors.Black, BorderWidth * 2) {DashPattern = new[] {5d, 5d}})
                .Add(new Paragraph().Add("test", Styles.TimesNewRoman10));
            Assert(nameof(DashBorders), document.CreatePng().Item1);
        }

        [Fact]
        public void SpanBackgroundColor()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1)                
            }));
            section.Add(new Paragraph()
                .Add(new Span("Choose ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("interfaces", TimesNewRoman10Bold))
                .Add(new Span(" over ", Styles.TimesNewRoman10))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                .Add(new Span("know something", new Font(TimesNewRoman, 18, XFontStyle.BoldItalic, PdfOptions))
                    .Brush(XBrushes.Red))
                .Add(new Span(" is going to be a baseclass, your first choice should be to make it an", Styles.TimesNewRoman10))
                .Add(new Span(" interface", TimesNewRoman10Bold))
                .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                    "variables should you change to an ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" class.", Styles.TimesNewRoman10)));
            Assert(nameof(SpanBackgroundColor), document.CreatePng().Item1);
        }

        [Fact]
        public void TextIndent()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(3),
                RightMargin = Cm(1.5)                
            }));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add(new Span("Choose ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("interfaces", TimesNewRoman10Bold))
                .Add(new Span(" over ", Styles.TimesNewRoman10))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                .Add(new Span("know something", new Font(TimesNewRoman, 18, XFontStyle.BoldItalic, PdfOptions))
                    .Brush(XBrushes.Red))
                .Add(new Span(" is going to be a baseclass, your first choice should be to make it an", Styles.TimesNewRoman10))
                .Add(new Span(" interface", TimesNewRoman10Bold))
                .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                    "variables should you change to an ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" class.", Styles.TimesNewRoman10)));
            section.Add(new Paragraph().TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Center).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Right).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            Assert(nameof(TextIndent), document.CreatePng().Item1);
        }

        [Fact]
        public void HorizontalAlign_Justify()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(3),
                RightMargin = Cm(1.5)                
            }));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
                .Add(new Span("Choose ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("interfaces", TimesNewRoman10Bold))
                .Add(new Span(" over ", Styles.TimesNewRoman10))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                .Add(new Span("know something", new Font(TimesNewRoman, 18, XFontStyle.BoldItalic, PdfOptions))
                    .Brush(XBrushes.Red))
                .Add(new Span(" is going to be a baseclass, your first choice should be to make it an", Styles.TimesNewRoman10))
                .Add(new Span(" interface", TimesNewRoman10Bold))
                .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                    "variables should you change to an ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" class.", Styles.TimesNewRoman10)));
            Assert(nameof(HorizontalAlign_Justify), document.CreatePng().Item1);
        }

        [Fact]
        public void HorizontalAlign_Justify_Underline()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                TopMargin = Cm(2),
                BottomMargin = Cm(2),
                LeftMargin = Cm(2.5),
                RightMargin = Cm(1.5)                
            }));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
                .Add(new Span("Choose composition first when creating new classes from existing classes. Only if " +
		            "inheritance is required by your design should it be used. If you use inheritance where " +
		            "composition will work, your designs will become needlessly complicated. ",
                    new Font(TimesNewRoman, 10, XFontStyle.Underline, PdfOptions))));
	        section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
		        .Add(new Span("Choose composition first when creating new classes from existing classes. Only if " +
			        "inheritance is required by your design should it be used. If you use inheritance where " +
			        "composition will work, your designs will become needlessly complicated. ",
			        new Font(TimesNewRoman, 20, XFontStyle.Underline, PdfOptions))));
	        section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
		        .Add(new Span("Choose composition first when creating new classes from existing classes. Only if " +
			        "inheritance is required by your design should it be used. If you use inheritance where " +
			        "composition will work, your designs will become needlessly complicated. ",
			        new Font(TimesNewRoman, 10, XFontStyle.Underline | XFontStyle.Bold, PdfOptions))));
	        section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
		        .Add(new Span("Choose composition first when creating new classes from existing classes. Only if " +
			        "inheritance is required by your design should it be used. If you use inheritance where " +
			        "composition will work, your designs will become needlessly complicated. ",
			        new Font(TimesNewRoman, 20, XFontStyle.Underline | XFontStyle.Bold, PdfOptions))));
            Assert(nameof(HorizontalAlign_Justify_Underline), document.CreatePng().Item1);
        }

        [Fact]
        public void ParagraphLineSpacingFunc()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            section.Add(new Paragraph()
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    Styles.TimesNewRoman10));
            section.Add(new Paragraph()
                .LineSpacingFunc(_ => _ * 1.5)
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if ",
                    Styles.TimesNewRoman10)
                .Add(new Span("inheritance", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(" is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    Styles.TimesNewRoman10));
            section.Add(new Paragraph()
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    Styles.TimesNewRoman10));
            Assert(nameof(ParagraphLineSpacingFunc), document.CreatePng().Item1);
        }

        [Fact]
        public void Test1()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                LeftMargin = XUnit.FromCentimeter(3),
                RightMargin = XUnit.FromCentimeter(1.5),
                TopMargin = XUnit.FromCentimeter(0),
                BottomMargin = XUnit.FromCentimeter(0),
            }));
            Table1(section);
            Table1(section);
            Table2(section);
            Table1(section);
            Table1(section);
            Table1(section);
            Table1(section);
            Table1(section);
            Table2(section);
            Table1(section);
            Table1(section);
            Assert(nameof(Test1), document.CreatePng().Item1);
        }

        [Fact]
        public void Test2()
        {
            //var pageSettings = new PageSettings {
            //    PageHeight = Px(650),
            //    LeftMargin = XUnit.FromCentimeter(3),
            //    RightMargin = XUnit.FromCentimeter(1.5),
            //};
            //var tables = new [] {
            //    Table1(pageSettings),
            //    Table3(pageSettings),
            //    Table1(pageSettings),
            //};
            //Assert(nameof(Test2), CreatePng(pageSettings, tables));
        }

        [Fact]
        public void Test3()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings {
                LeftMargin = XUnit.FromCentimeter(3),
                RightMargin = XUnit.FromCentimeter(1.5)
            }));
            Table5(section);
            Table6(section);
            Table4(section);
            Table2(section);
            Table1(section);
            Assert(nameof(Test3), document.CreatePng().Item1);
        }

        [Fact]
        public void PaymentOrderTest()
        {
            var document = new Document();
            PaymentOrder.AddSection(document, new PaymentData {
                IncomingDate = new DateTime(2018, 1, 23),
                OutcomingDate = new DateTime(2018, 1, 23),
            });
            Assert(nameof(PaymentOrderTest), document.CreatePng().Item1);
        }

        [Fact]
        public void SvoTest()
        {
            var document = new Document();
            Svo.AddSection(document,new SvoData());
            Assert(nameof(SvoTest), document.CreatePng().Item1);
        }

        [Fact]
        public void ContractDealPassportTest()
        {
            var document = new Document();
            ContractDealPassport.AddSection(document, new ContractDealPassportData());
            Assert(nameof(ContractDealPassportTest), document.CreatePng().Item1);
        }

        [Fact]
        public void LoanAgreementDealPassportTest()
        {
            var document = new Document();
            LoanAgreementDealPassport.AddSection(document, new LoanAgreementDealPassportData());
            Assert(nameof(LoanAgreementDealPassportTest), document.CreatePng().Item1);
        }

        [Fact]
        public void BackgroundColor()
        {
            var document = new Document();
            var pageSettings = new PageSettings {
                TopMargin = Cm(2),
                BottomMargin = Cm(2),
                LeftMargin = Cm(2),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            var table = section.AddTable();
            var c1 = table.AddColumn(Cm(1));
            var c2 = table.AddColumn(Cm(1));
            var r1 = table.AddRow().Height(Cm(1));
            r1[c1].Colspan(c2).Rowspan(2).BackgroundColor(XColors.LightGray)
                .Add(TimesNewRoman10("123456789012345678901234567890123456789012345678901234567890"));
            table.AddRow().Height(Cm(1));
            var r3 = table.AddRow().Height(Cm(1));
            r3[c1].Border(Top, BorderWidth * 10)
                .Add(TimesNewRoman10("1"));
            r3[c2].Border(Top, BorderWidth)
                .Add(TimesNewRoman10("2"));
            Assert(nameof(BackgroundColor), document.CreatePng().Item1);
        }

        [Fact]
        public void Intersection_of_border_lines_in_left_column()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var table = section.AddTable();
            var c1 = table.AddColumn(Cm(1));
            var r1 = table.AddRow().Height(Cm(1));
            var r2 = table.AddRow().Height(Cm(1));
            r1[c1].Border(Bottom, Cm(0.1));
            r2[c1].Border(Left, Cm(0.1));
            Assert(nameof(Intersection_of_border_lines_in_left_column), document.CreatePng().Item1);
        }

        [Fact]
        public void FontFamiliesTest()
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings()));
            var type = typeof(FontFamilies);
            foreach (var propertyInfo in type.GetProperties(BindingFlags.Static | BindingFlags.Public))
            {
                var fontFamilyInfo = (FontFamilyInfo) propertyInfo.GetValue(null);
                var text = fontFamilyInfo.OriginalName;
                var size = 12;
                section.Add(new Paragraph()
                    .Add($"***{text}***", new Font(DefaultFontFamilies.Roboto, size, XFontStyle.Regular, PdfOptions)));
                section.Add(new Paragraph()
                    .Add(text, new Font(fontFamilyInfo, size, XFontStyle.Regular, PdfOptions)));
                section.Add(new Paragraph()
                    .Add(text, new Font(fontFamilyInfo, size, XFontStyle.Bold, PdfOptions)));
                section.Add(new Paragraph()
                    .Add(text, new Font(fontFamilyInfo, size, XFontStyle.Italic, PdfOptions)));
                section.Add(new Paragraph()
                    .Add(text, new Font(fontFamilyInfo, size, XFontStyle.BoldItalic, PdfOptions)));
                section.Add(new Paragraph()
                    .Add("", new Font(DefaultFontFamilies.Roboto, size, XFontStyle.Regular, PdfOptions)));
            }
            Xunit.Assert.NotNull(document.CreatePdf());
        }

        [Fact]
        public void FontResolverTest()
        {
            var singleFontFamilies = new[]
                {
                    Wingdings,
                    Wingdings2,
                }
                .Select(_ => _.Name).ToHashSet();
            var type = typeof(FontFamilies);
            foreach (var propertyInfo in type.GetProperties(BindingFlags.Static | BindingFlags.Public))
            {
                var fontFamilyInfo = (FontFamilyInfo) propertyInfo.GetValue(null);
                if (singleFontFamilies.Contains(fontFamilyInfo.Name))
                    SingleFont(fontFamilyInfo);
                else
                    MultipleFonts(fontFamilyInfo);
            }
        }

        private static void MultipleFonts(FontFamilyInfo familyInfo)
        {
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, true, false);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, false, true);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, true, true);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, false, false);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
        }
        
        private static void SingleFont(FontFamilyInfo familyInfo)
        {
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, true, false);
                Xunit.Assert.True(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, false, true);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.True(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, true, true);
                Xunit.Assert.True(info.MustSimulateBold);
                Xunit.Assert.True(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
            {
                var info = GlobalFontSettings.FontResolver.ResolveTypeface(familyInfo.Name, false, false);
                Xunit.Assert.False(info.MustSimulateBold);
                Xunit.Assert.False(info.MustSimulateItalic);
                Xunit.Assert.NotEmpty(GlobalFontSettings.FontResolver.GetFont(info.FaceName));
                Xunit.Assert.NotEqual(DefaultFaceName, info.FaceName);
            }
        }
        
        private static string DefaultFaceName
        {
            get
            {
                var fontFamilyInfo = DefaultFontFamilies.Roboto;
                return GlobalFontSettings.FontResolver.ResolveTypeface(fontFamilyInfo.Name, false, false).FaceName;
            }
        }

        public static Table Table(PageSettings pageSettings)
        {
            var table = new Table();
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454));
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            var r1 = table.AddRow();
            {
                var cell = r1[c1];
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма прописью"));
            }
            {
                var cell = r1[c2];
                cell.Colspan(c5);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                paragraph.Alignment(HorizontalAlign.Center);
                paragraph.Margin(Left, Px(10 * 10 * 5));
                cell.Add(paragraph);
            }
            var r2 = table.AddRow();
            {
                var cell = r2[c1];
                cell.Colspan(c2);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("ИНН"));
            }
            {
                var cell = r2[c3];
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("КПП"));
            }
            {
                var cell = r2[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма"));
            }
            {
                var cell = r2[c5];
                cell.Rowspan(2);
                cell.Border(Top, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10("777-33");
                paragraph.Alignment(HorizontalAlign.Center);
                cell.Add(paragraph);
            }
            var r3 = table.AddRow();
            r3.Height(Px(100));
            {
                var cell = r3[c1];
                cell.Colspan(c3);
                cell.Rowspan(2);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4 * 5)));
                paragraph.Alignment(HorizontalAlign.Center);
                paragraph.Margin(Top, Px(10));
                paragraph.Margin(Left, Px(60));
                paragraph.Margin(Right, Px(30));
                cell.Add(paragraph);
                cell.Border(Right, BorderWidth);
            }
            var r4 = table.AddRow();
            r4.Height(Px(100));
            {
                var cell = r4[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сч. №"));
            }
            var r5 = table.AddRow();
            {
                var cell = r5[c1];
                cell.Colspan(c3);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Плательщик"));
            }
            r5[c4].Border(Bottom, BorderWidth);
            return table;
        }

        private static void Table1(Section section)
        {
            var pageSettings = section.PageSettings;
            var table = section.AddTable();
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454));
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            var r1 = table.AddRow();
            {
                var cell = r1[c1];
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма прописью"));
            }
            {
                var cell = r1[c2];
                cell.Colspan(c5);
                cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1))));
            }
            var r2 = table.AddRow();
            {
                var cell = r2[c1];
                cell.Colspan(c2);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("ИНН"));
            }
            {
                var cell = r2[c3];
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("КПП"));
            }
            {
                var cell = r2[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма"));
            }
            {
                var cell = r2[c5];
                cell.Rowspan(2);
                cell.Border(Top, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Add(TimesNewRoman10("777-33"));
            }
            var r3 = table.AddRow();
            r3.Height(Px(100));
            {
                var cell = r3[c1];
                cell.Colspan(c3);
                cell.Rowspan(2);
                cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4 * 5))));
                cell.Border(Right, BorderWidth);
            }
            var r4 = table.AddRow();
            r4.Height(Px(100));
            {
                var cell = r4[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сч. №"));
            }
            var r5 = table.AddRow();
            {
                var cell = r5[c1];
                cell.Colspan(c3);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Плательщик"));
            }
            r5[c4].Border(Bottom, BorderWidth);
        }

        private static void Table2(Section section)
        {
            var pageSettings = section.PageSettings;
            var table = section.AddTable();
            var c0 = table.AddColumn(Px(202));
            var c1 = table.AddColumn(Px(257));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(257));
            var c4 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin - BorderWidth
                - table.Columns.Sum(_ => _.Width));
            for (var i = 0; i < 101; i++)
            {
                var row = table.AddRow();
                {
                    var cell = row[c0];
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Left, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10($"№ {i}"));
                }
                {
                    var cell = row[c1];
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Колонка 2"));
                }
                {
                    var cell = row[c2];
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Колонка 3"));
                }
                {
                    var cell = row[c3];
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Колонка 4"));
                }
                {
                    var cell = row[c4];
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Колонка 5"));
                }
            }
        }

        public static Table Table3(PageSettings pageSettings)
        {
            var table = new Table();
            var ИНН1 = table.AddColumn(Px(202));
            var ИНН2 = table.AddColumn(Px(257));
            var КПП = table.AddColumn(Px(454));
            var сумма = table.AddColumn(Px(144));
            var суммаValue = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(ИНН2);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("ИНН"));
                }
                {
                    var cell = row[КПП];
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("КПП"));
                }
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("Сумма"));
                }
                {
                    var cell = row[суммаValue];
                    cell.Rowspan(2);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Add(TimesNewRoman10("777-33"));
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Rowspan(2);
                    cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*25))));
                    cell.Border(Right, BorderWidth);
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Сч. №"));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Плательщик"));
                }
                row[сумма].Border(Bottom, BorderWidth);
            }
            return table;
        }

        private static void Table4(Section section)
        {
            var pageSettings = section.PageSettings;
            var table = section.AddTable();
            var ИНН1 = table.AddColumn(Px(202));
            var ИНН2 = table.AddColumn(Px(257));
            var КПП = table.AddColumn(Px(454));
            var сумма = table.AddColumn(Px(144));
            var суммаValue = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Сумма прописью"));
                }
                {
                    var cell = row[ИНН2];
                    cell.Colspan(суммаValue);
                    cell.Add(TimesNewRoman60Bold(string.Join(" ", Enumerable.Repeat("Сто рублей", 1))));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(ИНН2);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("ИНН"));
                }
                {
                    var cell = row[КПП];
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("КПП"));
                }
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("Сумма"));
                }
                {
                    var cell = row[суммаValue];
                    cell.Rowspan(2);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Add(TimesNewRoman60Bold("777-33"));
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Rowspan(2);
                    cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5))));
                    cell.Border(Right, BorderWidth);
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Сч. №"));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Плательщик"));
                }
                row[сумма].Border(Bottom, BorderWidth);
            }
        }

        private static void Table5(Section section)
        {
            var pageSettings = section.PageSettings;
            var table = section.AddTable();
            var ИНН1 = table.AddColumn(Px(202));
            var ИНН2 = table.AddColumn(Px(257));
            var КПП = table.AddColumn(Px(454));
            var сумма = table.AddColumn(Px(144));
            var суммаValue = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10(@"a

aaaaaaaaa ")
                        .Add(new Span("0123", new Font(Arial, 12, XFontStyle.Bold, PdfOptions)))
                        .Add(new Span("у", Styles.TimesNewRoman10))
                        .Add(new Span("567", new Font(Arial, 12, XFontStyle.Bold, PdfOptions)))
                        .Add(new Span("ЙЙЙ", Styles.TimesNewRoman10)));
                }
                {
                    var cell = row[ИНН2];
                    cell.Colspan(суммаValue);
                    cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1))));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(ИНН2);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("ИНН"));
                }
                {
                    var cell = row[КПП];
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("КПП"));
                }
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Top, BorderWidth);
                    cell.Add(TimesNewRoman10("Сумма"));
                }
                {
                    var cell = row[суммаValue];
                    cell.Rowspan(2);
                    cell.Border(Top, BorderWidth);
                    cell.Border(Bottom, BorderWidth);
                    cell.Add(TimesNewRoman10("777-33"));
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Rowspan(2);
                    cell.Add(TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5))));
                    cell.Border(Right, BorderWidth);
                }
            }
            {
                var row = table.AddRow();
                row.Height(Px(100));
                {
                    var cell = row[сумма];
                    cell.Rowspan(2);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Сч. №"));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.Border(Bottom, BorderWidth);
                    cell.Border(Right, BorderWidth);
                    cell.Add(TimesNewRoman10("Плательщик"));
                }
                row[сумма].Border(Bottom, BorderWidth);
            }
        }

        private static void Table6(Section section)
        {
            var pageSettings = section.PageSettings;
            var table = section.AddTable();
            var c1 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin);
            var r1 = table.AddRow();
            {
                r1[c1].Add(new Paragraph().Add(new Span(
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated.",
                    Styles.TimesNewRoman10)));
            }
            var r2 = table.AddRow();
            {
                r2[c1].Add(new Paragraph().Add(new Span(
                    "Choose interfaces over abstract classes. If you know something is going to be a base" +
                    "class, your first choice should be to make it an interface, and only if you’re forced to" +
                    "have method definitions or member variables should you change to an abstract class.",
                    Styles.TimesNewRoman10)));
            }
            var r3 = table.AddRow();
            {
                r3[c1].Add(new Paragraph()
                    .Add(new Span("Choose ", Styles.TimesNewRoman10))
                    .Add(new Span("interfaces", TimesNewRoman10Bold))
                    .Add(new Span(" over ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" classes. If you know something is going to be a baseclass, your" +
                        " first choice should be to make it an ", Styles.TimesNewRoman10))
                    .Add(new Span("interface", TimesNewRoman10Bold))
                    .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                        "variables should you change to an ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" class.", Styles.TimesNewRoman10)));
            }
            var r4 = table.AddRow();
            {
                r4[c1].Add(new Paragraph()
                    .Add(new Span("Choose ", Styles.TimesNewRoman10))
                    .Add(new Span("interfaces", TimesNewRoman10Bold))
                    .Add(new Span(" over ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                    .Add(new Span("know something", new Font(TimesNewRoman, 18, XFontStyle.BoldItalic, PdfOptions)))
                    .Add(new Span(" is going to be a baseclass, your first choice should be to make it an ",
                        Styles.TimesNewRoman10))
                    .Add(new Span("interface", TimesNewRoman10Bold))
                    .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                        "variables should you change to an ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" class.", Styles.TimesNewRoman10)));
            }
            var r5 = table.AddRow();
            {
                r5[c1].Add(new Paragraph()
                    .Add(new Span("Choose ", Styles.TimesNewRoman10))
                    .Add(new Span("interfaces", TimesNewRoman10Bold))
                    .Add(new Span(" over ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                    .Add(new Span("know something", new Font(TimesNewRoman, 18, XFontStyle.BoldItalic, PdfOptions))
                        .Brush(XBrushes.Red))
                    .Add(new Span(" is going to be a baseclass, your first choice should be to make it an ",
                        Styles.TimesNewRoman10))
                    .Add(new Span("interface", TimesNewRoman10Bold))
                    .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                        "variables should you change to an ", Styles.TimesNewRoman10))
                    .Add(new Span("abstract", TimesNewRoman10Bold))
                    .Add(new Span(" class.", Styles.TimesNewRoman10)));
            }
        }

        public static Table Table7(PageSettings pageSettings)
        {
            var table = new Table();
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454));
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            var r1 = table.AddRow();
            {
                var cell = r1[c1];
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма прописью"));
            }
            {
                var cell = r1[c2];
                cell.Colspan(c5);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                paragraph.Alignment(HorizontalAlign.Center);
                paragraph.Margin(Left, Px(10 * 10 * 5));
                cell.Add(paragraph);
            }
            var r2 = table.AddRow();
            {
                var cell = r2[c1];
                cell.Colspan(c2);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("ИНН"));
            }
            {
                var cell = r2[c3];
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("КПП"));
            }
            {
                var cell = r2[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Top, BorderWidth);
                cell.Add(TimesNewRoman10("Сумма"));
            }
            {
                var cell = r2[c5];
                cell.Rowspan(2);
                cell.Border(Top, BorderWidth);
                cell.Border(Bottom, BorderWidth);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10("777-33");
                paragraph.Alignment(HorizontalAlign.Center);
                cell.Add(paragraph);
            }
            var r3 = table.AddRow();
            r3.Height(Px(100));
            {
                var cell = r3[c1];
                cell.Colspan(c3);
                cell.Rowspan(2);
                cell.VerticalAlign(VerticalAlign.Center);
                var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4 * 5)));
                paragraph.Alignment(HorizontalAlign.Center);
                cell.Add(paragraph);
                cell.Border(Right, BorderWidth);
            }
            var r4 = table.AddRow();
            r4.Height(Px(100));
            {
                var cell = r4[c4];
                cell.Rowspan(2);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Сч. №"));
            }
            var r5 = table.AddRow();
            {
                var cell = r5[c1];
                cell.Colspan(c3);
                cell.Border(Bottom, BorderWidth);
                cell.Border(Right, BorderWidth);
                cell.Add(TimesNewRoman10("Плательщик"));
            }
            r5[c4].Border(Bottom, BorderWidth);
            return table;
        }

        public static Paragraph TimesNewRoman10(string text) => new Paragraph().Add(text, Styles.TimesNewRoman10);

        public static Paragraph TimesNewRoman60Bold(string text) => new Paragraph().Add(text, Styles.TimesNewRoman60Bold);

        private static void Assert(string folderName, List<byte[]> pages)
        {
            for (var index = 0; index < pages.Count; index++)
            {
                var path = GetPageFileName(folderName, index);
                if (!File.ReadAllBytes(path).SequenceEqual(pages[index]))
                {
                    var tempFileName = Path.ChangeExtension(Path.GetTempFileName(), ".png");
                    ImageUtil.ImageDiff(path, pages[index], tempFileName);
                    throw new Exception($@"Images are different, see file:///{tempFileName.Replace(@"\", "/")}");
                }
            }
        }

        // ReSharper disable once UnusedMember.Local
        private static void SavePages(string folderName, List<byte[]> pages)
        {
            for (var index = 0; index < pages.Count; index++)
                File.WriteAllBytes(
                    GetPageFileName(folderName, index),
                    pages[index]);
        }

        private static string GetPageFileName(string folderName, int index)
        {
            return Path.Combine(
                Path.Combine(
                    Path.Combine(GetPath(), "TestData"),
                    folderName),
                $"Page{index + 1}.png");
        }

        public static string GetPath([CallerFilePath] string path = "") => new FileInfo(path).Directory.FullName;
    }

    public class ImageInfo
    {
        public byte[] Bytes { get; }
        public double PointWidth { get; }
        public double PointHeight { get; }

        public ImageInfo(byte[] bytes, double pointWidth, double pointHeight)
        {
            Bytes = bytes;
            PointWidth = pointWidth;
            PointHeight = pointHeight;
        }
    }
}