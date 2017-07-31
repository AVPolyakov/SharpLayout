using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Xunit;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public class Tests
    {
        [Fact]
        public void Test1()
        {
            var pageSettings = new PageSettings {
                LeftMargin = XUnit.FromCentimeter(3),
                RightMargin = XUnit.FromCentimeter(1.5),
                TopMargin = XUnit.FromCentimeter(0),
                BottomMargin = XUnit.FromCentimeter(0),
            };
            var tables = new [] {
                Table1(pageSettings),
                Table1(pageSettings),
                Table2(pageSettings),
                Table1(pageSettings),
                Table1(pageSettings),
                Table1(pageSettings),
                Table1(pageSettings),
                Table1(pageSettings),
                Table2(pageSettings),
                Table1(pageSettings),
                Table1(pageSettings),
            };
            Assert(nameof(Test1), CreatePng(pageSettings, tables));
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
            var pageSettings = new PageSettings {
                LeftMargin = XUnit.FromCentimeter(3),
                RightMargin = XUnit.FromCentimeter(1.5),
                IsHighlightCells = false
            };
            var tables = new[] {
                Table5(pageSettings),
                Table6(pageSettings),
                Table4(pageSettings),
                Table2(pageSettings),
                Table1(pageSettings),
            };
            Assert(nameof(Test3), CreatePng(pageSettings, tables));
        }

        [Fact]
        public void PaymentOrderTest()
        {
            PageSettings pageSettings;
            List<Table> tables;
            PaymentOrder.GetContent(out pageSettings, out tables, false);
            Assert(nameof(PaymentOrderTest), CreatePng(pageSettings, tables));
        }

        [Fact]
        public void Test4()
        {
            //var pageSettings = new PageSettings {
            //    LeftMargin = XUnit.FromCentimeter(3),
            //    RightMargin = XUnit.FromCentimeter(1.5),
            //    PageHeight = Px(700),
            //    IsHighlightCells = true
            //};
            //var tables = new [] {
            //    Table7(pageSettings),
            //};
            //Assert(nameof(Test4), CreatePng(pageSettings, tables));
        }

        public static Table Table(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin); 
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454)); 
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var r1 = table.AddRow();
                {
                    var cell = r1[c1];
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма прописью");
                }
                {
                    var cell = r1[c2];
                    cell.Colspan(c5);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                    paragraph.Alignment = ParagraphAlignment.Center;
                    paragraph.LeftMargin = Px(10 * 10 * 5);
                    cell.Paragraph = paragraph;
                }
            }
            {
                var r2 = table.AddRow();
                {
                    var cell = r2[c1];
                    cell.Colspan(c2);
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = r2[c3];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = r2[c4];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = r2[c5];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("777-33");
                    paragraph.Alignment = ParagraphAlignment.Center;
                    cell.Paragraph = paragraph;
                }
            }
            {
                var r3 = table.AddRow();
                r3.Height = Px(100);
                {
                    var cell = r3[c1];
                    cell.Colspan(c3);
                    cell.MergeDown = 1;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5)));
                    paragraph.Alignment = ParagraphAlignment.Center;
                    paragraph.TopMargin = Px(10);
                    paragraph.LeftMargin = Px(60);
                    paragraph.RightMargin = Px(30);
                    cell.Paragraph = paragraph;
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var r4 = table.AddRow();
                r4.Height = Px(100);
                {
                    var cell = r4[c4];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var r5 = table.AddRow();
                {
                    var cell = r5[c1];
                    cell.Colspan(c3);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                r5[c4].BottomBorder = BorderWidth;
            }
            return table;
        }

        public static Table Table1(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454));
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var r1 = table.AddRow();
                {
                    var cell = r1[c1];
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма прописью");
                }
                {
                    var cell = r1[c2];
                    cell.Colspan(c5);
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                }
            }
            {
                var r2 = table.AddRow();
                {
                    var cell = r2[c1];
                    cell.Colspan(c2);
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = r2[c3];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = r2[c4];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = r2[c5];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("777-33");
                }
            }
            {
                var r3 = table.AddRow();
                r3.Height = Px(100);
                {
                    var cell = r3[c1];
                    cell.Colspan(c3);
                    cell.MergeDown = 1;
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5)));
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var r4 = table.AddRow();
                r4.Height = Px(100);
                {
                    var cell = r4[c4];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var r5 = table.AddRow();
                {
                    var cell = r5[c1];
                    cell.Colspan(c3);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                r5[c4].BottomBorder = BorderWidth;
            }
            return table;
        }

        public static Table Table2(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
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
                    cell.BottomBorder = BorderWidth;
                    cell.LeftBorder = BorderWidth;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10($"№ {i}");
                }
                {
                    var cell = row[c1];
                    cell.BottomBorder = BorderWidth;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Колонка 2");
                }
                {
                    var cell = row[c2];
                    cell.BottomBorder = BorderWidth;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Колонка 3");
                }
                {
                    var cell = row[c3];
                    cell.BottomBorder = BorderWidth;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Колонка 4");
                }
                {
                    var cell = row[c4];
                    cell.BottomBorder = BorderWidth;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Колонка 5");
                }
            }
            return table;
        }

        public static Table Table3(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
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
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = row[КПП];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = row[суммаValue];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("777-33");
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.MergeDown = 1;
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*25)));
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                row[сумма].BottomBorder = BorderWidth;
            }
            return table;
        }

        public static Table Table4(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
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
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма прописью");
                }
                {
                    var cell = row[ИНН2];
                    cell.Colspan(суммаValue);
                    cell.Paragraph = TimesNewRoman60(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(ИНН2);
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = row[КПП];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = row[суммаValue];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman60("777-33");
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.MergeDown = 1;
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5)));
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                row[сумма].BottomBorder = BorderWidth;
            }
            return table;
        }

        public static Table Table5(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
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
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10(@"a

aaaaaaaaa ")
                        .Add(new Span("0123", new XFont("Arial", 12, XFontStyle.Bold, pdfOptions)))
                        .Add(new Span("у", TimesNewRoman10Font))
                        .Add(new Span("567", new XFont("Arial", 12, XFontStyle.Bold, pdfOptions)))
                        .Add(new Span("ЙЙЙ", TimesNewRoman10Font));
                }
                {
                    var cell = row[ИНН2];
                    cell.Colspan(суммаValue);
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(ИНН2);
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = row[КПП];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = row[суммаValue];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("777-33");
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.MergeDown = 1;
                    cell.Paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5)));
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var row = table.AddRow();
                row.Height = Px(100);
                {
                    var cell = row[сумма];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var row = table.AddRow();
                {
                    var cell = row[ИНН1];
                    cell.Colspan(КПП);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                row[сумма].BottomBorder = BorderWidth;
            }
            return table;
        }

        public static Table Table6(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
            var c1 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin);
            {
                var r1 = table.AddRow();
                {
                    r1[c1].Paragraph = new Paragraph().Add(new Span(
                        "Choose composition first when creating new classes from existing classes. Only if " +
                        "inheritance is required by your design should it be used. If you use inheritance where " +
                        "composition will work, your designs will become needlessly complicated.",
                        TimesNewRoman10Font));
                }
            }
            {
                var r2 = table.AddRow();
                {
                    r2[c1].Paragraph = new Paragraph().Add(new Span(
                        "Choose interfaces over abstract classes. If you know something is going to be a base" +
                        "class, your first choice should be to make it an interface, and only if you’re forced to" +
                        "have method definitions or member variables should you change to an abstract class.",
                        TimesNewRoman10Font));
                }
            }
            {
                var r3 = table.AddRow();
                {
                    r3[c1].Paragraph = new Paragraph()
                        .Add(new Span("Choose ", TimesNewRoman10Font))
                        .Add(new Span("interfaces", TimesNewRoman10BoldFont))
                        .Add(new Span(" over ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" classes. If you know something is going to be a baseclass, your" +
                            " first choice should be to make it an ", TimesNewRoman10Font))
                        .Add(new Span("interface", TimesNewRoman10BoldFont))
                        .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                            "variables should you change to an ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" class.", TimesNewRoman10Font));
                }
            }
            {
                var r4 = table.AddRow();
                {
                    r4[c1].Paragraph = new Paragraph()
                        .Add(new Span("Choose ", TimesNewRoman10Font))
                        .Add(new Span("interfaces", TimesNewRoman10BoldFont))
                        .Add(new Span(" over ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" classes. If you ", TimesNewRoman10Font))
                        .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, pdfOptions)))
                        .Add(new Span(" is going to be a baseclass, your first choice should be to make it an ",
                            TimesNewRoman10Font))
                        .Add(new Span("interface", TimesNewRoman10BoldFont))
                        .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                            "variables should you change to an ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" class.", TimesNewRoman10Font));
                }
            }
            {
                var r5 = table.AddRow();
                {
                    r5[c1].Paragraph = new Paragraph()
                        .Add(new Span("Choose ", TimesNewRoman10Font))
                        .Add(new Span("interfaces", TimesNewRoman10BoldFont))
                        .Add(new Span(" over ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" classes. If you ", TimesNewRoman10Font))
                        .Add(new Span("know something", new XFont("Times New Roman", 18, XFontStyle.BoldItalic, pdfOptions)) {
                            Brush = XBrushes.Red
                        })
                        .Add(new Span(" is going to be a baseclass, your first choice should be to make it an ",
                            TimesNewRoman10Font))
                        .Add(new Span("interface", TimesNewRoman10BoldFont))
                        .Add(new Span(", and only if you’re forced tohave method definitions or member " +
                            "variables should you change to an ", TimesNewRoman10Font))
                        .Add(new Span("abstract", TimesNewRoman10BoldFont))
                        .Add(new Span(" class.", TimesNewRoman10Font));
                }
            }
            return table;
        }

        public static Table Table7(PageSettings pageSettings)
        {
            var table = new Table(pageSettings.LeftMargin);
            var c1 = table.AddColumn(Px(202));
            var c2 = table.AddColumn(Px(257));
            var c3 = table.AddColumn(Px(454));
            var c4 = table.AddColumn(Px(144));
            var c5 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                - table.Columns.Sum(_ => _.Width));
            {
                var r1 = table.AddRow();
                {
                    var cell = r1[c1];
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма прописью");
                }
                {
                    var cell = r1[c2];
                    cell.Colspan(c5);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Сто рублей", 1)));
                    paragraph.Alignment = ParagraphAlignment.Center;
                    paragraph.LeftMargin = Px(10 * 10 * 5);
                    cell.Paragraph = paragraph;
                }
            }
            {
                var r2 = table.AddRow();
                {
                    var cell = r2[c1];
                    cell.Colspan(c2);
                    cell.RightBorder = cell.TopBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("ИНН");
                }
                {
                    var cell = r2[c3];
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("КПП");
                }
                {
                    var cell = r2[c4];
                    cell.MergeDown = 1;
                    cell.TopBorder = cell.BottomBorder = cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сумма");
                }
                {
                    var cell = r2[c5];
                    cell.MergeDown = 1;
                    cell.BottomBorder = cell.TopBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("777-33");
                    paragraph.Alignment = ParagraphAlignment.Center;
                    cell.Paragraph = paragraph;
                }
            }
            {
                var r3 = table.AddRow();
                r3.Height = Px(100);
                {
                    var cell = r3[c1];
                    cell.Colspan(c3);
                    cell.MergeDown = 1;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10(string.Join(" ", Enumerable.Repeat("Ромашка", 4*5)));
                    paragraph.Alignment = ParagraphAlignment.Center;
                    cell.Paragraph = paragraph;
                    cell.RightBorder = BorderWidth;
                }
            }
            {
                var r4 = table.AddRow();
                r4.Height = Px(100);
                {
                    var cell = r4[c4];
                    cell.MergeDown = 1;
                    cell.RightBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Сч. №");
                }
            }
            {
                var r5 = table.AddRow();
                {
                    var cell = r5[c1];
                    cell.Colspan(c3);
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Paragraph = TimesNewRoman10("Плательщик");
                }
                r5[c4].BottomBorder = BorderWidth;
            }
            return table;
        }

        private static readonly XPdfFontOptions pdfOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);

        public static Paragraph TimesNewRoman12Bold(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman12BoldFont));

        public static Paragraph TimesNewRoman10(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman10Font));

        public static Paragraph TimesNewRoman10_5(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman10_5Font));

        public static Paragraph TimesNewRoman9_5(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman9_5Font));

        public static Paragraph TimesNewRoman9(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman9Font));

        public static Paragraph TimesNewRoman11_5Bold(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman11_5BoldFont));

        public static Paragraph TimesNewRoman8(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman8Font));

        public static Paragraph TimesNewRoman60(string text) =>
            new Paragraph().Add(new Span(text, TimesNewRoman60BoldFont));

        public static readonly XFont TimesNewRoman12BoldFont = new XFont("Times New Roman", 12, XFontStyle.Bold, pdfOptions);

        public static readonly XFont TimesNewRoman10Font = new XFont("Times New Roman", 10, XFontStyle.Regular, pdfOptions);

        public static readonly XFont TimesNewRoman8Font = new XFont("Times New Roman", 8, XFontStyle.Regular, pdfOptions);

        public static readonly XFont TimesNewRoman10_5Font = new XFont("Times New Roman", 10.5, XFontStyle.Regular, pdfOptions);

        public static readonly XFont TimesNewRoman9_5Font = new XFont("Times New Roman", 9.5, XFontStyle.Regular, pdfOptions);

        public static readonly XFont TimesNewRoman9Font = new XFont("Times New Roman", 9, XFontStyle.Regular, pdfOptions);

        public static readonly XFont TimesNewRoman11_5BoldFont = new XFont("Times New Roman", 11.5, XFontStyle.Bold, pdfOptions);

        public static readonly XFont TimesNewRoman10BoldFont = new XFont("Times New Roman", 10, XFontStyle.Bold, pdfOptions);

        public static readonly XFont TimesNewRoman60BoldFont = new XFont("Times New Roman", 60, XFontStyle.Bold, pdfOptions);

        private static void Assert(string folderName, List<byte[]> pages)
        {
            for (var index = 0; index < pages.Count; index++)
                Xunit.Assert.True(File.ReadAllBytes(GetPageFileName(folderName, index))
                    .SequenceEqual(pages[index]));
        }

        public static List<byte[]> CreatePng(PageSettings pageSettings, IEnumerable<Table> tables)
        {
            var pages = new List<byte[]> {null};
            FillBitmap(xGraphics => TableRenderer.Draw(xGraphics, pageSettings,
                    (pageIndex, action) => FillBitmap(action, bitmap => pages.Add(ToBytes(bitmap)), pageSettings), tables),
                bitmap => pages[0] = ToBytes(bitmap),
                pageSettings);
            return pages;
        }

        private static void FillBitmap(Action<XGraphics> action, Action<Bitmap> action2, PageSettings pageSettings)
        {
            const int resolution = 254;
            var horzPixels = (int) (new XUnit(pageSettings.PageWidth).Inch * resolution);
            var vertPixels = (int) (new XUnit(pageSettings.PageHeight).Inch * resolution);
            using (var bitmap = new Bitmap(horzPixels, vertPixels))
            {
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, horzPixels, vertPixels);
                    using (var xGraphics = XGraphics.FromGraphics(graphics, new XSize(horzPixels, vertPixels)))
                    {
                        xGraphics.ScaleTransform(resolution / 72d);
                        action(xGraphics);
                    }
                }
                bitmap.SetResolution(resolution, resolution);
                action2(bitmap);
            }
        }

        private static byte[] ToBytes(Bitmap bitmap)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
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

        public const double BorderWidth = 0.5D;
    }
}