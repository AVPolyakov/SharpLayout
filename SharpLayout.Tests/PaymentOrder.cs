using static SharpLayout.Direction;
using static SharpLayout.Tests.Tests;
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
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(351));
                table.AddColumn(Px(125));
                var c3 = table.AddColumn(Px(351));
                var c4 = table.AddColumn();
                var c5 = table.AddColumn(Px(150));
                c4.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c1].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("23.01.2018").Alignment(HorizontalAlign.Center).DefaultMargin());
                r1[c3].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("23.01.2018").Alignment(HorizontalAlign.Center).DefaultMargin());
                r1[c5].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("0401060").Alignment(HorizontalAlign.Center).DefaultMargin());
                var r2 = table.AddRow();
                r2[c1].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman8("Поступ. в банк плат.").Alignment(HorizontalAlign.Center).DefaultMargin());
                r2[c3].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman8("Списано со сч. плат.").Alignment(HorizontalAlign.Center).DefaultMargin());
                table.AddRow().Height(Px(40));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(900));
                var c2 = table.AddColumn(Px(352));
                table.AddColumn(Px(49));
                var c4 = table.AddColumn(Px(351));
                var c5 = table.AddColumn();
                var c6 = table.AddColumn(Px(70));
                c5.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(63));
                r1[c1].VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman12Bold("ПЛАТЕЖНОЕ ПОРУЧЕНИЕ № 17").DefaultMargin());
                r1[c2].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("17.01.2018").Alignment(HorizontalAlign.Center).DefaultMargin());
                r1[c4].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("Электронно").Alignment(HorizontalAlign.Center).DefaultMargin());
                r1[c6].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("02").Alignment(HorizontalAlign.Center).DefaultMargin());
                var r2 = table.AddRow();
                r2[c2].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman8("Дата").Alignment(HorizontalAlign.Center).DefaultMargin());
                r2[c4].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman8("Вид платежа").Alignment(HorizontalAlign.Center).DefaultMargin());
                table.AddRow().Height(Px(33));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(202));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(144));
                r1[c1].Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сумма прописью").DefaultMargin());
                r1[c2].Border(Bottom)
                    .Add(TimesNewRoman10("Одна тысяча семьсот семьдесят рублей 00 копеек").LeftIndentMargin());
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(502));
                var c2 = table.AddColumn(Px(501));
                var c3 = table.AddColumn(Px(150));
                var c4 = table.AddColumn(Px(200));
                var c5 = table.AddColumn(Px(200));
                var c6 = table.AddColumn();
                c6.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(45));
                r1[c1].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("ИНН  0123456789").DefaultMargin());
                r1[c2].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("КПП  012345678").LeftIndentMargin());
                r1[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сумма").LeftIndentMargin());
                r1[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(TimesNewRoman10("1770-00").LeftIndentMargin());
                var r2 = table.AddRow().Height(Px(100));
                r2[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(TimesNewRoman10("Клиент012345678").DefaultMargin());
                var r3 = table.AddRow().Height(Px(102));
                r3[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сч. №").LeftIndentMargin());
                r3[c4].Colspan(c6).Rowspan(2)
                    .Add(TimesNewRoman10("01234567890123456789").LeftIndentMargin());
                var r4 = table.AddRow().Height(Px(45));
                r4[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Плательщик").DefaultMargin());
                var r5 = table.AddRow().Height(Px(48));
                r5[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(TimesNewRoman10("Клиент012345").DefaultMargin());
                r5[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("БИК").LeftIndentMargin());
                r5[c4].Colspan(c6).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("012345678").LeftIndentMargin());
                var r6 = table.AddRow().Height(Px(49));
                r6[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сч. №").LeftIndentMargin());
                r6[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(TimesNewRoman10("01234567890123456789").LeftIndentMargin());
                var r7 = table.AddRow().Height(Px(45));
                r7[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Банк плательщика").DefaultMargin());
                var r8 = table.AddRow().Height(Px(49));
                r8[c1].Colspan(c2).Rowspan(2).Border(Right)
                    .Add(TimesNewRoman10("ПАО \"ТЕСТБАНК\" Г.МОСКВА").DefaultMargin());
                r8[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("БИК").LeftIndentMargin());
                r8[c4].Colspan(c6).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("012345678").LeftIndentMargin());
                var r9 = table.AddRow().Height(Px(58));
                r9[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сч. №").LeftIndentMargin());
                r9[c4].Colspan(c6).Rowspan(2)
                    .Add(TimesNewRoman10("01234567890123456789").LeftIndentMargin());
                var r10 = table.AddRow().Height(Px(45));
                r10[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Банк получателя").DefaultMargin());
                var r11 = table.AddRow().Height(Px(45));
                r11[c1].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("ИНН  0123456789").DefaultMargin());
                r11[c2].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("КПП  01234567").LeftIndentMargin());
                r11[c3].Rowspan(2).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Сч. №").LeftIndentMargin());
                r11[c4].Colspan(c6).Rowspan(2).Border(Bottom)
                    .Add(TimesNewRoman10("01234567890123456789").LeftIndentMargin());
                var r12 = table.AddRow().Height(Px(98));
                r12[c1].Colspan(c2).Rowspan(3).Border(Right)
                    .Add(TimesNewRoman10("Клиент012345678").DefaultMargin());
                var r13 = table.AddRow().Height(Px(47));
                r13[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Вид оп.").LeftIndentMargin());
                r13[c4].Border(Right).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("01").LeftIndentMargin());
                r13[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Срок плат.").LeftIndentMargin());
                r13[c6].VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("2").LeftIndentMargin());
                var r14 = table.AddRow().Height(Px(47));
                r14[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Наз. пл.").LeftIndentMargin());
                r14[c4].Border(Right).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("2").LeftIndentMargin());
                r14[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Очер. плат.").LeftIndentMargin());
                r14[c6].VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("5").LeftIndentMargin());
                var r15 = table.AddRow().Height(Px(46));
                r15[c1].Colspan(c2).VerticalAlign(VerticalAlign.Bottom).Border(Right | Bottom)
                    .Add(TimesNewRoman10("Получатель").DefaultMargin());
                r15[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Код").LeftIndentMargin());
                r15[c4].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("2").LeftIndentMargin());
                r15[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman10("Рез. поле").LeftIndentMargin());
                r15[c6].Border(Bottom).VerticalAlign(VerticalAlign.Center);
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
                c7.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(45));
                r1[c1].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c2].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c3].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c4].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c5].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c6].Border(Right | Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c7].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("2").DefaultMargin().Alignment(HorizontalAlign.Center));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Px(246));
                r1[c1].Add(TimesNewRoman10(@"Тест
В том числе НДС 270,00").DefaultMargin());
                var r2 = table.AddRow().Height(Px(46));
                r2[c1].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("Назначение платежа").DefaultMargin());
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(603));
                var c2 = table.AddColumn(Px(601));
                var c3 = table.AddColumn();
                c3.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var height = Px(25);
                var r1 = table.AddRow().Height(Px(102) - height);
                r1[c2].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman10("Подписи").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c3].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman10("Отметки банка").DefaultMargin().Alignment(HorizontalAlign.Center));
                var r2 = table.AddRow().Height(height);
                r2[c3].Rowspan(4).Add(Stamp());
                var r3 = table.AddRow();
                r3[c2].Border(Bottom)
                    .Add(TimesNewRoman10("").DefaultMargin().Alignment(HorizontalAlign.Center));
                var r4 = table.AddRow().Height(Px(144));
                r4[c1].VerticalAlign(VerticalAlign.Top)
                    .Add(TimesNewRoman10("М.П.").DefaultMargin().Alignment(HorizontalAlign.Center));
                r4[c2].Border(Bottom).VerticalAlign(VerticalAlign.Bottom)
                    .Add(TimesNewRoman10("").DefaultMargin().Alignment(HorizontalAlign.Center));
                table.AddRow();
            }
        }

        private static Table Stamp()
        {
            var table = new Table {HorizontalAlign = HorizontalAlign.Right};
            var c1 = table.AddColumn(Px(520));
            var r1 = table.AddRow();
            r1[c1].Border(Left | Top | Right | Bottom)
                .Add(TimesNewRoman10(@"ХХХ ХХХХ ХХХХХХ ХХХХХ
XXX 012345678").DefaultMargin());
            var r2 = table.AddRow();
            r2[c1].Border(Left | Right | Bottom)
                .Add(TimesNewRoman10(@"Xxxxxxx ")
                    .Add(new Span(@"xxxxxxxxx", TimesNewRoman10BoldFont))
                    .Add(new Span(@"
ЭЛЕКТРОННО
xxx xxxxx Xxxxxx xxxxxxxxx
23.02.2018", TimesNewRoman10Font)).DefaultMargin());
            var r3 = table.AddRow();
            r3[c1].Border(Left | Right | Bottom)
                .Add(TimesNewRoman10(@"Xxxx xxxxx xxxxxxxx xxxxxx xxxxx xxxxx xxxxxxxxxx xxxxx xxxxxxxx")
                    .DefaultMargin());
            return table;
        }

        private static readonly double cellMargin = Cm(0.05);

        private static Paragraph DefaultMargin(this Paragraph @this) => @this.Margin(Left | Right, cellMargin);

        private static Paragraph LeftIndentMargin(this Paragraph @this) => @this.Margin(Left, cellMargin + Cm(0.1)).Margin(Right, cellMargin);
    }
}