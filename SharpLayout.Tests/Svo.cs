using PdfSharp;
using static SharpLayout.Direction;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class Svo
    {
        public static void AddSection(Document document)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(2),
                Orientation = PageOrientation.Landscape
            };
            var font = TimesNewRoman9_5;
            var headerFont = TimesNewRoman9;
            var section = document.Add(new Section(pageSettings));
            var rowHeight = Cm(0.49);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                table.AddRow()[c1].Add(Paragraph.Add("Код формы по ОКУД 0406009", TimesNewRoman10_5)
                    .Margin(Bottom, Px(11)).Alignment(HorizontalAlign.Right));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(6.3));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("Наименование уполномоченного банка", font));
                r1[c2].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font));
                var r2 = table.AddRow().Height(rowHeight);
                r2[c1].VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("Наименование резидента", font));
                r2[c2].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                table.AddRow()[c1].Add(new Paragraph().Add("СПРАВКА О ВАЛЮТНЫХ ОПЕРАЦИЯХ", TimesNewRoman11_5Bold)
                    .Alignment(HorizontalAlign.Center).Margin(Top, Px(41)));
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1146));
                var c2 = table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                table.AddRow()[c2].Add(new Paragraph().Add("от  ", TimesNewRoman11_5Bold).Margin(Top, Px(30 - 18)));
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1199));
                var c2 = table.AddColumn(Px(200));
                table.AddRow().Height(Px(28))[c2].TopBorder = BorderWidth;
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(8));
                var c2 = table.AddColumn(Px(405));
                var c3 = table.AddColumn();
                var c4 = table.AddColumn(Px(325));
                c3.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("Номер счета резидента в уполномоченном банке", font));
                r1[c2].Colspan(c4).Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font).Alignment(HorizontalAlign.Center));
                var r2 = table.AddRow().Height(rowHeight);
                r2[c1].Add(Paragraph.Add("Код страны банка-нерезидента", font)).VerticalAlign(VerticalAlign.Center);
                r2[c2].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font).Alignment(HorizontalAlign.Center));
                r2[c3].VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("Признак корректировки", font).Alignment(HorizontalAlign.Right).Margin(Right, Cm(0.2) + cellMargin));
                r2[c4].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font).Alignment(HorizontalAlign.Center));
            }
            {
                var table = section.AddTable();
                table.AddColumn(pageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                table.AddRow().Height(Px(35));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(1.05));
                var c2 = table.AddColumn();
                var c3 = table.AddColumn(Cm(2) - Px(1));
                var c4 = table.AddColumn(Cm(2) - Px(1));
                var c5 = table.AddColumn(Cm(2.1) - Px(1));
                var c6 = table.AddColumn(Px(211) - Px(1));
                var c7 = table.AddColumn(Cm(4.2) - Px(211) - Px(1));
                var c8 = table.AddColumn(Cm(3.05) - Px(1));
                var c9 = table.AddColumn(Px(225) - Px(1));
                var c10 = table.AddColumn(Cm(4.5) - Px(225) - Px(1));
                var c11 = table.AddColumn(Cm(2.2) - Px(1));
                var c12 = table.AddColumn(Cm(2.2) - Px(1));
                c2.Width = pageSettings.PageWidthWithoutMargins - BorderWidth - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c1].Rowspan(2).Border(Left | Top | Right | Bottom)
                    .Add(Paragraph.Add("№ п/п", headerFont));
                r1[c2].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Уведомление, распоряжение, расчетный или иной документ", headerFont));
                r1[c3].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Дата операции", headerFont));
                r1[c4].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Признак платежа", headerFont));
                r1[c5].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Код вида валютной операции", headerFont));
                r1[c6].Colspan(c7).Border(Top | Right | Bottom, BorderWidth
                    ).Add(Paragraph.Add("Сумма операции", headerFont));
                r1[c8].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add(@"Номер ПС
или номер и (или) дата договора (контракта)", headerFont));
                r1[c9].Colspan(c10).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add(@"Сумма операции
в единицах валюты контракта (кредитного договора)", headerFont));
                r1[c11].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Срок возврата аванса", headerFont));
                r1[c12].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(Paragraph.Add("Ожидаемый срок", headerFont));
                var r2 = table.AddRow();
                r2[c6].Border(Right | Bottom)
                    .Add(Paragraph.Add("код валюты", headerFont));
                r2[c7].Border(Right | Bottom)
                    .Add(Paragraph.Add("сумма", headerFont));
                r2[c9].Border(Right | Bottom)
                    .Add(Paragraph.Add("код валюты", headerFont));
                r2[c10].Border(Right | Bottom)
                    .Add(Paragraph.Add("сумма", headerFont));
                var r3 = table.AddRow();
                for (var i = 0; i < table.Columns.Count; i++)
                    r3[table.Columns[i]].Border(Right | Bottom)
                        .Add(Paragraph.Add($"{i + 1}", headerFont));
                r3[c1].Border(Left);
                foreach (var row in new[] {r1, r2, r3})
                foreach (var column in table.Columns)
                foreach (var paragraph in row[column].VerticalAlign(VerticalAlign.Center).Paragraphs)
                    paragraph.Alignment(HorizontalAlign.Center);
                for (var i = 0; i < 3; i++)
                {
                    var row = table.AddRow().Height(Cm(0.46));
                    row[c1].Border(Left)
                        .Add(Paragraph.Add($"{i + 1}", font));
                    row[c2].Add(Paragraph.Add("", font));
                    row[c3].Add(Paragraph.Add("", font));
                    row[c4].Add(Paragraph.Add("", font));
                    row[c5].Add(Paragraph.Add("", font));
                    row[c6].Add(Paragraph.Add("", font));
                    row[c7].Add(Paragraph.Add("", font));
                    row[c8].Add(Paragraph.Add("", font));
                    row[c9].Add(Paragraph.Add("", font));
                    row[c10].Add(Paragraph.Add("", font));
                    row[c11].Add(Paragraph.Add("", font));
                    row[c12].Add(Paragraph.Add("", font));
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center).Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(363));
                table.AddRow().Height(Px(41))[c1].Border(Bottom);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                table.AddRow()[c1].Add(new Paragraph().Add("Примечание.", font).Margin(Top, Px(1)).Margin(Bottom, Px(4)));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(255));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - BorderWidth - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Cm(0.46));
                r1[c1].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("№ строки", headerFont).Alignment(HorizontalAlign.Center));
                r1[c2].Border(Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("Содержание", headerFont).Alignment(HorizontalAlign.Center));
                for (var i = 0; i < 2; i++)
                {
                    var row = table.AddRow().Height(Cm(0.46));
                    row[c1].Border(Left)
                        .Add(Paragraph.Add("", font).Alignment(HorizontalAlign.Center));
                    row[c2].Add(Paragraph.Add("", font));
                    foreach (var column in table.Columns)
                        row[column].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                table.AddRow()[c1].Add(new Paragraph().Add("Информация уполномоченного банка", font).Margin(Top, Px(3)).Margin(Bottom, Px(4)));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidthWithoutMargins - BorderWidth;
                var r1 = table.AddRow().Height(Cm(0.46));
                r1[c1].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(Paragraph.Add("", font).Alignment(HorizontalAlign.Center));
            }
        }

        private static Paragraph Paragraph => new Paragraph().Margin(Left | Right, cellMargin);

        private static readonly double cellMargin = Cm(0.05);
    }
}