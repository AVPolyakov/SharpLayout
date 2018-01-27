using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using Xunit;
using static SharpLayout.Direction;
using static SharpLayout.Util;
using static SharpLayout.Tests.Styles;

namespace SharpLayout.Tests
{
    public class Tests
    {
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
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
                var footers = section.AddFooters();
                var c1 = footers.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = footers.AddRow().Height(Px(700));
                r1[c1].Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add(new Span(new PageNumber(), Styles.TimesNewRoman10))
                    .Add(" из ", Styles.TimesNewRoman10)
                    .Add(new Span(new PageCount(), Styles.TimesNewRoman10)));
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
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
                .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, PdfOptions)) {
                    Brush = XBrushes.Red
                })
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify).TextIndent(Cm(1))
                .Add(new Span("Choose ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("interfaces", TimesNewRoman10Bold))
                .Add(new Span(" over ", Styles.TimesNewRoman10))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, PdfOptions)) {
                    Brush = XBrushes.Red
                })
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Center).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Right).TextIndent(Cm(1))
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
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
                    new XFont("Times New Roman", 12, XFontStyle.Regular, PdfOptions)));
            section.Add(new Paragraph().Alignment(HorizontalAlign.Justify)
                .Add(new Span("Choose ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("interfaces", TimesNewRoman10Bold))
                .Add(new Span(" over ", Styles.TimesNewRoman10))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" classes. If you ", Styles.TimesNewRoman10))
                .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, PdfOptions)) {
                    Brush = XBrushes.Red
                })
                .Add(new Span(" is going to be a baseclass, your first choice should be to make it an", Styles.TimesNewRoman10))
                .Add(new Span(" interface", TimesNewRoman10Bold))
                .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                    "variables should you change to an ", Styles.TimesNewRoman10).BackgroundColor(XColors.LightGray))
                .Add(new Span("abstract", TimesNewRoman10Bold))
                .Add(new Span(" class.", Styles.TimesNewRoman10)));
            Assert(nameof(HorizontalAlign_Justify), document.CreatePng().Item1);
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
            PaymentOrder.AddSection(document);
            Assert(nameof(PaymentOrderTest), document.CreatePng().Item1);
        }

        [Fact]
        public void SvoTest()
        {
            var document = new Document();
            Svo.AddSection(document);
            Assert(nameof(SvoTest), document.CreatePng().Item1);
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
                paragraph.LeftMargin = Px(10 * 10 * 5);
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
                paragraph.TopMargin = Px(10);
                paragraph.LeftMargin = Px(60);
                paragraph.RightMargin = Px(30);
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

        public static void Table1(Section section)
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

        public static void Table2(Section section)
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

        public static void Table4(Section section)
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

        public static void Table5(Section section)
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
                        .Add(new Span("0123", new XFont("Arial", 12, XFontStyle.Bold, PdfOptions)))
                        .Add(new Span("у", Styles.TimesNewRoman10))
                        .Add(new Span("567", new XFont("Arial", 12, XFontStyle.Bold, PdfOptions)))
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

        public static void Table6(Section section)
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
                    .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, PdfOptions)))
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
                    .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, PdfOptions)) {
                        Brush = XBrushes.Red
                    })
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
                paragraph.LeftMargin = Px(10 * 10 * 5);
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
                Xunit.Assert.True(File.ReadAllBytes(GetPageFileName(folderName, index))
                    .SequenceEqual(pages[index]));
        }

        public static void SavePages(string folderName, List<byte[]> pages)
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
}