using static SharpLayout.Direction;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class PaymentClaim
    {
        public static void AddSection(Document document)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(0.8),
                LeftMargin = Cm(2),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            {
                var table = section.AddTable().Font(TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                var c2 = table.AddColumn(Cm(6));
                var c3 = table.AddColumn();
                var c4 = table.AddColumn(Cm(5.5));
                new[] {c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow();
                r1[c2].Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("Подписи"));
                r1[c4].Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("Отметки банка получателя"));
                var r2 = table.AddRow().Height(Cm(1.07));
                r2[c2].Border(Bottom);
                var r3 = table.AddRow().Height(Cm(1.43));
                r3[c2].Border(Bottom);
                r3[c1].Add(Paragraph.Alignment(HorizontalAlign.Center)
                    .Add("М.П."));
            }
            section.Add(new Paragraph().Add("", TimesNewRoman10));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                var c2 = table.AddColumn(Cm(5.5));
                new[] {c1}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow();
                r1[c1].Border(Top).LeftBottomTable();
                r1[c2].RightBottomTable();
            }
        }

        private static readonly double height = Cm(0.38 + 0.44);

        private static void LeftBottomTable(this Cell cell)
        {
            var table = new Table().Font(TimesNewRoman10);
            cell.Add(table);
            var c1 = table.AddColumn(Cm(1));
            var c2 = table.AddColumn(Cm(1.5));
            var c3 = table.AddColumn(Cm(2.5));
            var c4 = table.AddColumn();
            var c5 = table.AddColumn();
            var c6 = table.AddColumn(Cm(1.5));
            new[] {c4, c5}.Distribute(cell.Table.Columns[cell.ColumnIndex].Width - table.ColumnsWidth);
            var r1 = table.AddRow().Height(height);
            r1[c1].Border(Right | Bottom)
                .Add(Paragraph.Add("№ ч. плат."));
            r1[c2].Border(Right | Bottom)
                .Add(Paragraph.Add("№ плат. ордера"));
            r1[c3].Border(Right | Bottom)
                .Add(Paragraph.Add("Дата плат. ордера"));
            r1[c4].Border(Right | Bottom)
                .Add(Paragraph.Add("Сумма частичного платежа"));
            r1[c5].Border(Right | Bottom)
                .Add(Paragraph.Add("Сумма остатка платежа"));
            r1[c6].Border(Bottom)
                .Add(Paragraph.Add("Подпись"));
            foreach (var c in table.Columns)
            foreach (var p in r1[c].Paragraphs)
                p.Alignment(HorizontalAlign.Center);
            for (var i = 0; i < 1; i++)
            {
                var r = table.AddRow();
                r[c1].Border(Right)
                    .Add(Paragraph.Add(""));
                r[c2].Border(Right)
                    .Add(Paragraph.Add(""));
                r[c3].Border(Right)
                    .Add(Paragraph.Add(""));
                r[c4].Border(Right)
                    .Add(Paragraph.Add(""));
                r[c5].Border(Right)
                    .Add(Paragraph.Add(""));
                r[c6]
                    .Add(Paragraph.Add(""));
            }
        }

        private static void RightBottomTable(this Cell cell)
        {
            var table = new Table().Font(TimesNewRoman10);
            cell.Add(table);
            var c1 = table.AddColumn();
            new[] {c1}.Distribute(cell.Table.Columns[cell.ColumnIndex].Width - table.ColumnsWidth);
            var r1 = table.AddRow();
            r1[c1].Add(Paragraph.Alignment(HorizontalAlign.Center)
                .Add("Дата помещения в картотеку"));
            var r2 = table.AddRow().Height(height + BorderWidth);
            r2[c1].Add(Paragraph.Alignment(HorizontalAlign.Center)
                .Add("01.01.2018"));
            var r3 = table.AddRow();
            r3[c1].Add(Paragraph.Alignment(HorizontalAlign.Center)
                .Add("Отметки банка плательщика"));
            var r4 = table.AddRow();
            r4[c1].Add(Stamp());
        }

        private static Table Stamp()
        {
            var table = new Table().Font(TimesNewRoman10).Alignment(HorizontalAlign.Right).Border(BorderWidth);
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

        private static Paragraph Paragraph => new Paragraph().Margin(Left | Right, Cm(0.05));
    }
}