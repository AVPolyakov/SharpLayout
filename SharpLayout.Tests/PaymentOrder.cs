using System.Linq;
using static SharpLayout.Tests.Tests;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class PaymentOrder
    {
        public static void AddSection(Document document)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            var cellMargin = Cm(0.05);
            var leftIndent = Cm(0.1);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(351));
                table.AddColumn(Px(125));
                var c3 = table.AddColumn(Px(351));
                var c4 = table.AddColumn();
                var c5 = table.AddColumn(Px(150));
                c4.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                {
                    var cell = r1[c1];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("23.01.2018");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c3];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("23.01.2018");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c5];
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("0401060");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                {
                    var cell = r2[c1];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman8("Поступ. в банк плат.");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r2[c3];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman8("Списано со сч. плат.");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r3 = table.AddRow();
                r3.Height = Px(40);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(900));
                var c2 = table.AddColumn(Px(352));
                table.AddColumn(Px(49));
                var c4 = table.AddColumn(Px(351));
                var c5 = table.AddColumn();
                var c6 = table.AddColumn(Px(70));
                c5.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Px(63);
                {
                    var cell = r1[c1];
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman12Bold("ПЛАТЕЖНОЕ ПОРУЧЕНИЕ № 17");
                    paragraph.LeftMargin = Px(1);
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c2];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("17.01.2018");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c4];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("Электронно");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c6];
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("02");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                {
                    var cell = r2[c2];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman8("Дата");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r2[c4];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman8("Вид платежа");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r3 = table.AddRow();
                r3.Height = Px(33);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(202));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Px(144);
                {
                    var cell = r1[c1];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сумма прописью");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c2];
                    cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Одна тысяча семьсот семьдесят рублей 00 копеек");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(502));
                var c2 = table.AddColumn(Px(501));
                var c3 = table.AddColumn(Px(150));
                var c4 = table.AddColumn(Px(200));
                var c5 = table.AddColumn(Px(200));
                var c6 = table.AddColumn();
                c6.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Px(45);
                {
                    var cell = r1[c1];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("ИНН  0123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c2];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("КПП  012345678");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c3];
                    cell.Rowspan = 2;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сумма");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c4];
                    cell.Colspan(c6);
                    cell.Rowspan = 2;
                    cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("1770-00");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                r2.Height = Px(100);
                {
                    var cell = r2[c1];
                    cell.Colspan(c2);
                    cell.Rowspan = 2;
                    cell.RightBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Клиент012345678");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r3 = table.AddRow();
                r3.Height = Px(102);
                {
                    var cell = r3[c3];
                    cell.Rowspan = 2;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сч. №");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r3[c4];
                    cell.Colspan(c6);
                    cell.Rowspan = 2;
                    var paragraph = TimesNewRoman10("01234567890123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r4 = table.AddRow();
                r4.Height = Px(45);
                {
                    var cell = r4[c1];
                    cell.Colspan(c2);
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Плательщик");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r5 = table.AddRow();
                r5.Height = Px(48);
                {
                    var cell = r5[c1];
                    cell.Colspan(c2);
                    cell.Rowspan = 2;
                    cell.RightBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Клиент012345");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r5[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("БИК");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r5[c4];
                    cell.Colspan(c6);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("012345678");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r6 = table.AddRow();
                r6.Height = Px(49);
                {
                    var cell = r6[c3];
                    cell.Rowspan = 2;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сч. №");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r6[c4];
                    cell.Colspan(c6);
                    cell.Rowspan = 2;
                    cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("01234567890123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r7 = table.AddRow();
                r7.Height = Px(45);
                {
                    var cell = r7[c1];
                    cell.Colspan(c2);
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Банк плательщика");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r8 = table.AddRow();
                r8.Height = Px(49);
                {
                    var cell = r8[c1];
                    cell.Colspan(c2);
                    cell.Rowspan = 2;
                    cell.RightBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("ПАО \"ТЕСТБАНК\" Г.МОСКВА");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r8[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("БИК");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r8[c4];
                    cell.Colspan(c6);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("012345678");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r9 = table.AddRow();
                r9.Height = Px(58);
                {
                    var cell = r9[c3];
                    cell.Rowspan = 2;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сч. №");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r9[c4];
                    cell.Colspan(c6);
                    cell.Rowspan = 2;
                    var paragraph = TimesNewRoman10("01234567890123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r10 = table.AddRow();
                r10.Height = Px(45);
                {
                    var cell = r10[c1];
                    cell.Colspan(c2);
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Банк получателя");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r11 = table.AddRow();
                r11.Height = Px(45);
                {
                    var cell = r11[c1];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("ИНН  0123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r11[c2];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("КПП  01234567");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r11[c3];
                    cell.Rowspan = 2;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Сч. №");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r11[c4];
                    cell.Colspan(c6);
                    cell.Rowspan = 2;
                    cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("01234567890123456789");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r12 = table.AddRow();
                r12.Height = Px(98);
                {
                    var cell = r12[c1];
                    cell.Colspan(c2);
                    cell.Rowspan = 3;
                    cell.RightBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Клиент012345678");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r13 = table.AddRow();
                r13.Height = Px(47);
                {
                    var cell = r13[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Вид оп.");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r13[c4];
                    cell.RightBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("01");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r13[c5];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Срок плат.");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r13[c6];
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r14 = table.AddRow();
                r14.Height = Px(47);
                {
                    var cell = r14[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Наз. пл.");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r14[c4];
                    cell.RightBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r14[c5];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Очер. плат.");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r14[c6];
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("5");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                var r15 = table.AddRow();
                r15.Height = Px(46);
                {
                    var cell = r15[c1];
                    cell.Colspan(c2);
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("Получатель");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                {
                    var cell = r15[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Код");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r15[c4];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r15[c5];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman10("Рез. поле");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.LeftMargin = paragraph.LeftMargin.ValueOrDefault() + leftIndent;
                    cell.Add(paragraph);
                }
                {
                    var cell = r15[c6];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(452));
                var c2 = table.AddColumn(Px(300));
                var c3 = table.AddColumn(Px(100));
                var c4 = table.AddColumn(Px(251));
                var c5 = table.AddColumn(Px(350));
                var c6 = table.AddColumn(Px(250));
                var c7 = table.AddColumn();
                c7.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Px(45);
                {
                    var cell = r1[c1];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c2];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c3];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c4];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c5];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c6];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c7];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("2");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Px(246);
                {
                    var cell = r1[c1];
                    var paragraph = TimesNewRoman10(@"Тест
В том числе НДС 270,00");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                r2.Height = Px(46);
                {
                    var cell = r2[c1];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("Назначение платежа");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(603));
                var c2 = table.AddColumn(Px(601));
                var c3 = table.AddColumn();
                c3.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin
                    - table.Columns.Sum(_ => _.Width);
                var height = Px(25);
                var r1 = table.AddRow();
                r1.Height = Px(102) - height;
                {
                    var cell = r1[c2];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman10("Подписи");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c3];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman10("Отметки банка");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                r2.Height = height;
                {
                    var cell = r2[c3];
                    cell.Rowspan = 4;
                    cell.Add(Stamp(cellMargin));
                }
                var r3 = table.AddRow();
                {
                    var cell = r3[c2];
                    cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman10("");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                var r4 = table.AddRow();
                r4.Height = Px(144);
                {
                    var cell = r4[c1];
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    var paragraph = TimesNewRoman10("М.П.");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r4[c2];
                    cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Bottom;
                    var paragraph = TimesNewRoman10("");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                table.AddRow();
            }
        }

        private static Table Stamp(double cellMargin)
        {
            var table = new Table {HorizontalAlignment = HorizontalAlignment.Right};
            var c1 = table.AddColumn(Px(520));
            var r1 = table.AddRow();
            {
                var cell = r1[c1];
                cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                var paragraph = TimesNewRoman10(@"ХХХ ХХХХ ХХХХХХ ХХХХХ
XXX 012345678");
                paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                cell.Add(paragraph);
            }
            var r2 = table.AddRow();
            {
                var cell = r2[c1];
                cell.LeftBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                var paragraph = TimesNewRoman10(@"Xxxxxxx ")
                    .Add(new Span(@"xxxxxxxxx", TimesNewRoman10BoldFont))
                    .Add(new Span(@"
ЭЛЕКТРОННО
xxx xxxxx Xxxxxx xxxxxxxxx
23.02.2018", TimesNewRoman10Font));
                paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                cell.Add(paragraph);
            }
            var r3 = table.AddRow();
            {
                var cell = r3[c1];
                cell.LeftBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                var paragraph = TimesNewRoman10(@"Xxxx xxxxx xxxxxxxx xxxxxx xxxxx xxxxx xxxxxxxxxx xxxxx xxxxxxxx");
                paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                cell.Add(paragraph);
            }
            return table;
        }
    }
}