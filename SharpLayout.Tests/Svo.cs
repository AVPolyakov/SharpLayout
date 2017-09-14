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
            var section = document.Add(new Section(pageSettings));
            var rowHeight = Cm(0.49);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(TimesNewRoman10_5("Код формы по ОКУД 0406009")
                    .DefaultMargin().Margin(Bottom, Px(11)).Alignment(HorizontalAlign.Right));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(6.3));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("Наименование уполномоченного банка").DefaultMargin());
                r1[c2].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").DefaultMargin());
                var r2 = table.AddRow().Height(rowHeight);
                r2[c1].VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("Наименование резидента").DefaultMargin());
                r2[c2].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").DefaultMargin());
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(TimesNewRoman11_5Bold("СПРАВКА О ВАЛЮТНЫХ ОПЕРАЦИЯХ").Alignment(HorizontalAlign.Center).Margin(Top, Px(41)));
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1146));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c2].Add(TimesNewRoman11_5Bold("от  ").Margin(Top, Px(30 - 18)));
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1199));
                var c2 = table.AddColumn(Px(200));
                var r1 = table.AddRow().Height(Px(28));
                r1[c2].TopBorder = BorderWidth;
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
                    .Add(TimesNewRoman9_5("Номер счета резидента в уполномоченном банке").DefaultMargin());
                r1[c2].Colspan(c4).Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").Alignment(HorizontalAlign.Center).DefaultMargin());
                var r2 = table.AddRow().Height(rowHeight);
                r2[c1].Add(TimesNewRoman9_5("Код страны банка-нерезидента").DefaultMargin()).VerticalAlign(VerticalAlign.Center);
                r2[c2].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").Alignment(HorizontalAlign.Center).DefaultMargin());
                r2[c3].VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("Признак корректировки").Alignment(HorizontalAlign.Right).Margin(Left, cellMargin).Margin(Right, Cm(0.2) + cellMargin));
                r2[c4].Border(Left | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").Alignment(HorizontalAlign.Center).DefaultMargin());
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
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
                    .Add(TimesNewRoman9("№ п/п"));
                r1[c2].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Уведомление, распоряжение, расчетный или иной документ"));
                r1[c3].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Дата операции"));
                r1[c4].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Признак платежа"));
                r1[c5].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Код вида валютной операции"));
                r1[c6].Colspan(c7).Border(Top | Right | Bottom, BorderWidth
                    ).Add(TimesNewRoman9("Сумма операции"));
                r1[c8].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9(@"Номер ПС
или номер и (или) дата договора (контракта)"));
                r1[c9].Colspan(c10).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9(@"Сумма операции
в единицах валюты контракта (кредитного договора)"));
                r1[c11].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Срок возврата аванса"));
                r1[c12].Rowspan(2).Border(Top | Right | Bottom)
                    .Add(TimesNewRoman9("Ожидаемый срок"));
                var r2 = table.AddRow();
                r2[c6].Border(Right | Bottom)
                    .Add(TimesNewRoman9("код валюты"));
                r2[c7].Border(Right | Bottom)
                    .Add(TimesNewRoman9("сумма"));
                r2[c9].Border(Right | Bottom)
                    .Add(TimesNewRoman9("код валюты"));
                r2[c10].Border(Right | Bottom)
                    .Add(TimesNewRoman9("сумма"));
                var r3 = table.AddRow();
                for (var i = 0; i < table.Columns.Count; i++)
                    r3[table.Columns[i]].Border(Right | Bottom)
                        .Add(TimesNewRoman9($"{i + 1}"));
                r3[c1].Border(Left);
                foreach (var row in new[] {r1, r2, r3})
                foreach (var column in table.Columns)
                foreach (var paragraph in row[column].VerticalAlign(VerticalAlign.Center).Paragraphs)
                    paragraph.DefaultMargin().Alignment(HorizontalAlign.Center);
                for (var i = 0; i < 3; i++)
                {
                    var row = table.AddRow().Height(Cm(0.46));
                    row[c1].Border(Left)
                        .Add(TimesNewRoman9_5($"{i + 1}"));
                    row[c2].Add(TimesNewRoman9_5(""));
                    row[c3].Add(TimesNewRoman9_5(""));
                    row[c4].Add(TimesNewRoman9_5(""));
                    row[c5].Add(TimesNewRoman9_5(""));
                    row[c6].Add(TimesNewRoman9_5(""));
                    row[c7].Add(TimesNewRoman9_5(""));
                    row[c8].Add(TimesNewRoman9_5(""));
                    row[c9].Add(TimesNewRoman9_5(""));
                    row[c10].Add(TimesNewRoman9_5(""));
                    row[c11].Add(TimesNewRoman9_5(""));
                    row[c12].Add(TimesNewRoman9_5(""));
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center).Paragraphs)
                        paragraph.DefaultMargin().Alignment(HorizontalAlign.Center);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(363));
                var r1 = table.AddRow().Height(Px(41));
                r1[c1].Border(Bottom);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(TimesNewRoman9_5("Примечание.").Margin(Top, Px(1)).Margin(Bottom, Px(4)));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(255));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidthWithoutMargins - BorderWidth - table.ColumnsWidth;
                var r1 = table.AddRow().Height(Cm(0.46));
                r1[c1].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9("№ строки").DefaultMargin().Alignment(HorizontalAlign.Center));
                r1[c2].Border(Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9("Содержание").DefaultMargin().Alignment(HorizontalAlign.Center));
                for (var i = 0; i < 2; i++)
                {
                    var row = table.AddRow().Height(Cm(0.46));
                    row[c1].Border(Left)
                        .Add(TimesNewRoman9_5("").DefaultMargin().Alignment(HorizontalAlign.Center));
                    row[c2].Add(TimesNewRoman9_5("").DefaultMargin());
                    foreach (var column in table.Columns)
                        row[column].Border(Right | Bottom).VerticalAlign(VerticalAlign.Center);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(TimesNewRoman9_5("Информация уполномоченного банка").Margin(Top, Px(3)).Margin(Bottom, Px(4)));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidthWithoutMargins - BorderWidth;
                var r1 = table.AddRow().Height(Cm(0.46));
                r1[c1].Border(Left | Top | Right | Bottom).VerticalAlign(VerticalAlign.Center)
                    .Add(TimesNewRoman9_5("").DefaultMargin().Alignment(HorizontalAlign.Center));
            }
        }

        private static readonly double cellMargin = Cm(0.05);

        private static Paragraph DefaultMargin(this Paragraph paragraph) => paragraph.Margin(Left | Right, cellMargin);
    }
}