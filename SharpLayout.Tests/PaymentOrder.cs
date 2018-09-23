using System;
using PdfSharp.Drawing;
using static SharpLayout.Direction;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class PaymentOrder
    {
        public static void AddSection(Document document, PaymentData data)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(Px(351));
                table.AddColumn(Px(125));
                var c3 = table.AddColumn(Px(351));
                var c4 = table.AddColumn();
                var c5 = table.AddColumn(Px(150));
                c4.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c1].Border(Bottom).VerticalAlign(VerticalAlign.Bottom).Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add(() => data.IncomingDate, FormatDate));
                r1[c3].Border(Bottom).VerticalAlign(VerticalAlign.Bottom).Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add(() => data.OutcomingDate, FormatDate));
                r1[c5].Border(All).VerticalAlign(VerticalAlign.Center).Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("0401060"));
                var r2 = table.AddRow();
                r2[c1].VerticalAlign(VerticalAlign.Top).Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("Поступ. в банк плат.", TimesNewRoman8));
                r2[c3].VerticalAlign(VerticalAlign.Top).Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("Списано со сч. плат.", TimesNewRoman8));
                table.AddRow().Height(Px(40));
            }
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(Px(900));
                var c2 = table.AddColumn(Px(352));
                table.AddColumn(Px(49));
                var c4 = table.AddColumn(Px(351));
                var c5 = table.AddColumn();
                var c6 = table.AddColumn(Px(70));
                c5.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(63));
                r1[c1].VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("ПЛАТЕЖНОЕ ПОРУЧЕНИЕ № 17", TimesNewRoman12Bold));
                r1[c2].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("17.01.2018").Alignment(HorizontalAlign.Center));
                r1[c4].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("Электронно").Alignment(HorizontalAlign.Center));
                r1[c6].Border(All).VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("02").Alignment(HorizontalAlign.Center));
                var r2 = table.AddRow();
                r2[c2].VerticalAlign(VerticalAlign.Top)
                    .Add(Paragraph.Add("Дата", TimesNewRoman8).Alignment(HorizontalAlign.Center));
                r2[c4].VerticalAlign(VerticalAlign.Top)
                    .Add(Paragraph.Add("Вид платежа", TimesNewRoman8).Alignment(HorizontalAlign.Center));
                table.AddRow().Height(Px(33));
            }
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(Px(202));
                var c2 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow().Height(Px(144));
                r1[c1].Border(Right | Bottom)
                    .Add(Paragraph.Add("Сумма прописью"));
                r1[c2].Border(Bottom)
                    .Add(LeftIndentParagraph.Add("Одна тысяча семьсот семьдесят рублей 00 копеек"));
            }
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(Px(502));
                var c2 = table.AddColumn(Px(501));
                var c3 = table.AddColumn(Px(150));
                var c4 = table.AddColumn(Px(200));
                var c5 = table.AddColumn(Px(200));
                var c6 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow().Height(Px(45));
                r1[c1].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("ИНН  0123456789"));
                r1[c2].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("КПП  012345678"));
                r1[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(LeftIndentParagraph.Add("Сумма"));
                r1[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(LeftIndentParagraph.Add("1770-00"));
                var r2 = table.AddRow().Height(Px(100));
                r2[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(Paragraph.Add("Клиент012345678"));
                var r3 = table.AddRow().Height(Px(102));
                r3[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(LeftIndentParagraph.Add("Сч. №"));
                r3[c4].Colspan(c6).Rowspan(2)
                    .Add(LeftIndentParagraph.Add("01234567890123456789"));
                var r4 = table.AddRow().Height(Px(45));
                r4[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(Paragraph.Add("Плательщик"));
                var r5 = table.AddRow().Height(Px(48));
                r5[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(Paragraph.Add("Клиент012345"));
                r5[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("БИК"));
                r5[c4].Colspan(c6).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("012345678"));
                var r6 = table.AddRow().Height(Px(49));
                r6[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(LeftIndentParagraph.Add("Сч. №"));
                r6[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(LeftIndentParagraph.Add("01234567890123456789"));
                var r7 = table.AddRow().Height(Px(45));
                r7[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(Paragraph.Add("Банк плательщика"));
                var r8 = table.AddRow().Height(Px(49));
                r8[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(Paragraph.Add("ПАО \"ТЕСТБАНК\" Г.МОСКВА"));
                r8[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("БИК"));
                r8[c4].Colspan(c6).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("012345678"));
                var r9 = table.AddRow().Height(Px(58));
                r9[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(LeftIndentParagraph.Add("Сч. №"));
                r9[c4].Colspan(c6).Rowspan(2)
                    .Add(LeftIndentParagraph.Add("01234567890123456789"));
                var r10 = table.AddRow().Height(Px(45));
                r10[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(Paragraph.Add("Банк получателя"));
                var r11 = table.AddRow().Height(Px(45));
                r11[c1].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("ИНН  0123456789"));
                r11[c2].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("КПП  01234567"));
                r11[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(LeftIndentParagraph.Add("Сч. №"));
                r11[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(LeftIndentParagraph.Add("01234567890123456789"));
                var r12 = table.AddRow().Height(Px(98));
                r12[c1].Colspan(c2).Rowspan(3).Border(Right)
                    .Add(Paragraph.Add("Клиент012345678"));
                var r13 = table.AddRow().Height(Px(47));
                r13[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Вид оп."));
                r13[c4].Border(Right).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("01"));
                r13[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Срок плат."));
                r13[c6].VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("2"));
                var r14 = table.AddRow().Height(Px(47));
                r14[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Наз. пл."));
                r14[c4].Border(Right).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("2"));
                r14[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Очер. плат."));
                r14[c6].VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("5"));
                var r15 = table.AddRow().Height(Px(46));
                r15[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(Paragraph.Add("Получатель"));
                r15[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Код"));
                r15[c4].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("2"));
                r15[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(LeftIndentParagraph.Add("Рез. поле"));
                r15[c6].Border(Bottom).VerticalAlign(VerticalAlign.Center);
            }
            {
                var table = section.AddTable().Font(font).ContentVerticalAlign(VerticalAlign.Bottom)
                    .ContentAlign(HorizontalAlign.Center);
                var c1 = table.AddColumn(Px(452));
                var c2 = table.AddColumn(Px(300));
                var c3 = table.AddColumn(Px(100));
                var c4 = table.AddColumn(Px(251));
                var c5 = table.AddColumn(Px(350));
                var c6 = table.AddColumn(Px(250));
                var c7 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow().Height(Px(45));
                r1[c1].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c2].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c3].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c4].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c5].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c6].Border(Right | Bottom)
                    .Add(Paragraph.Add("2"));
                r1[c7].Border(Bottom)
                    .Add(Paragraph.Add("2"));
            }
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow().Height(Px(246));
                r1[c1].Add(Paragraph.Add(@"Тест
В том числе НДС 270,00"));
                var r2 = table.AddRow().Height(Px(46));
                r2[c1].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("Назначение платежа"));
            }
            {
                var table = section.AddTable().Font(font);
                var c1 = table.AddColumn(Px(603));
                var c2 = table.AddColumn(Px(601));
                var c3 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var height = Px(25);
                var r1 = table.AddRow().Height(Px(102) - height);
                r1[c2].VerticalAlign(VerticalAlign.Top)
                    .Add(Paragraph.Add("Подписи").Alignment(HorizontalAlign.Center));
                r1[c3].VerticalAlign(VerticalAlign.Top)
                    .Add(Paragraph.Add("Отметки банка").Alignment(HorizontalAlign.Center));
                var r2 = table.AddRow().Height(height);
                r2[c3].Rowspan(4).Add(Stamp());
                var r3 = table.AddRow();
                r3[c2].Border(Bottom)
                    .Add(Paragraph.Add("").Alignment(HorizontalAlign.Center));
                var r4 = table.AddRow().Height(Px(144));
                r4[c1].VerticalAlign(VerticalAlign.Top)
                    .Add(Paragraph.Add("М.П.").Alignment(HorizontalAlign.Center));
                r4[c2].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(Paragraph.Add("").Alignment(HorizontalAlign.Center));
                table.AddRow();
            }
        }        

        private static Table Stamp()
        {
            var table = new Table().Font(font).Alignment(HorizontalAlign.Right).Border(BorderWidth);
            var c1 = table.AddColumn(Px(520));
            var r1 = table.AddRow();
            r1[c1].Add(Paragraph.Add(@"ХХХ ХХХХ ХХХХХХ ХХХХХ
XXX 012345678"));
            var r2 = table.AddRow();
            r2[c1].Add(Paragraph.Add(@"Xxxxxxx ")
                    .Add(new Span(@"xxxxxxxxx", TimesNewRoman10Bold))
                    .Add(new Span(@"
ЭЛЕКТРОННО
xxx xxxxx Xxxxxx xxxxxxxxx
23.02.2018")));
            var r3 = table.AddRow();
            r3[c1].Add(Paragraph.Add(@"Xxxx xxxxx xxxxxxxx xxxxxx xxxxx xxxxx xxxxxxxxxx xxxxx xxxxxxxx"));
            return table;
        }

        private static readonly double cellMargin = Cm(0.05);
        private static Font font => TimesNewRoman10;
        private static Paragraph Paragraph => new Paragraph().Margin(Left | Right, cellMargin);
        private static Paragraph LeftIndentParagraph => Paragraph.Margin(Left, cellMargin + Cm(0.1));
    }

    public class PaymentData
    {
        public DateTime IncomingDate { get; set; }
        public DateTime OutcomingDate { get; set; }
    }
}