using static SharpLayout.Direction;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Util;

namespace SharpLayout.Tests
{
    public static class ContractDealPassport
    {
        public static void AddSection(Document document)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(1),
                BottomMargin = Cm(1),
                LeftMargin = Cm(3),
                RightMargin = Cm(1.5)
            };
            var section = document.Add(new Section(pageSettings));
            var verticalSpacing = 3d;
            var bigVerticalSpacing = verticalSpacing * 2;
            var smallVerticalSpacing = 2d;
            section.Add(new Paragraph().Margin(Bottom, verticalSpacing).Alignment(HorizontalAlign.Right)
                .Add("Код формы по ОКУД 0406005", TimesNewRoman10));
            section.Add(new Paragraph().Margin(Bottom, verticalSpacing).Alignment(HorizontalAlign.Right)
                .Add("Форма 1", TimesNewRoman11));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
	            var r1 = table.AddRow().Height(Cm(0.63));
                r1[c1].Border(All).VerticalAlign(VerticalAlign.Center)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                        .Add("Тест", TimesNewRoman11_5));
            }
            section.Add(new Paragraph().Margin(Bottom, verticalSpacing * 2).Alignment(HorizontalAlign.Center)
                .Add("Наименование банка ПС", TimesNewRoman9));
            var cellWidth = Cm(0.4);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(343));
                var c2 = table.AddColumn(Px(57));
                var c3 = table.AddColumn(Cm(2.2));
                var c4 = table.AddColumn(Px(65));
                var cellCount = 22;
                for (var i = 0; i < cellCount; i++)
                    table.AddColumn(cellWidth);
                new[] {c1, c4}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
	            var r1 = table.AddRow().Height(Cm(0.55));
                r1[c1].VerticalAlign(VerticalAlign.Bottom).Add(NormalParagraph
                    .Add("Паспорт сделки", TimesNewRoman13Bold));
                r1[c2].VerticalAlign(VerticalAlign.Bottom).Add(NormalParagraph.Alignment(HorizontalAlign.Right)
                    .Add("от", TimesNewRoman12));
                r1[c3].Border(Bottom).VerticalAlign(VerticalAlign.Bottom).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                    .Add("88.88.8888", TimesNewRoman12Bold));
                r1[c4].Border(Right).VerticalAlign(VerticalAlign.Bottom)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Right).Margin(Right, Px(10))
                        .Add("№", TimesNewRoman12));
                for (var i = 0; i < cellCount; i++)
                    r1[c4.Index + 1 + i].VerticalAlign(VerticalAlign.Bottom).Border(Top | Right | Bottom)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("12345678/1234/1234/1/1".CellSubstring(i, cellCount), TimesNewRoman12Bold));
            }
            section.Add(new Paragraph().Margin(Top, verticalSpacing * 2).Margin(Bottom, verticalSpacing)
                .Add("1. Сведения о резиденте", TimesNewRoman11_5Bold));
            {
                var table = section.AddTable().Font(TimesNewRoman9_5).ContentVerticalAlign(VerticalAlign.Center);
                var width1 = Px(270d + 10 + 3);
                var c1 = table.AddColumn(Cm(5.4) - width1);
                var c2 = table.AddColumn(width1);
                var width2 = Px(144);
                var c3 = table.AddColumn(Px(240) - width2);
                var c4 = table.AddColumn(width2);
                var c5 = table.AddColumn(Px(295));
                var c6 = table.AddColumn(Px(115));
                var c7 = table.AddColumn(Px(275));
                var c8 = table.AddColumn(Px(145));
                new[] {c4, c6, c8}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                var r1 = table.AddRow().Height(Cm(0.65));
	            r1[c1].Colspan(c2)
                    .Add(NormalParagraph.Add("1.1. Наименование"));
                r1[c3].Colspan(c8).Border(All)
                    .Add(NormalParagraph.Add(""));
                var r2 = table.AddRow();
                r2[c1]
                    .Add(NormalParagraph.Add("1.2. Адрес:"));
                r2[c2].Colspan(c4)
                    .Add(NormalParagraph.Add("Субъект Российской Федерации"));
                r2[c5].Colspan(c8).Border(Bottom | Left | Right)
                    .Add(NormalParagraph.Add(""));
                var r3 = table.AddRow();
                r3[c2].Colspan(c4)
                    .Add(NormalParagraph.Add("Район"));
                r3[c5].Colspan(c8).Border(Bottom | Left | Right)
                    .Add(NormalParagraph.Add(""));
                var r4 = table.AddRow();
                r4[c2].Colspan(c4)
                    .Add(NormalParagraph.Add("Город"));
                r4[c5].Colspan(c8).Border(Bottom | Left | Right)
                    .Add(NormalParagraph.Add(""));
                var r5 = table.AddRow();
                r5[c2].Colspan(c4)
                    .Add(NormalParagraph.Add("Населенный пункт"));
                r5[c5].Colspan(c8).Border(Bottom | Left | Right)
                    .Add(NormalParagraph.Add(""));
                var r6 = table.AddRow();
                r6[c2].Colspan(c4)
                    .Add(NormalParagraph.Add("Улица (проспект, переулок и т.д.)"));
                r6[c5].Colspan(c8).Border(Bottom | Left | Right)
                    .Add(NormalParagraph.Add(""));
                var r7 = table.AddRow();
                r7[c2].Colspan(c3)
                    .Add(NormalParagraph.Add("Номер дома (владение)"));
                r7[c4].Border(All)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Center).Add(""));
                r7[c5].Border(Right)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Right)
                        .Add("Корпус (строение)"));
                r7[c6].Border(Bottom | Right)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Center).Add(""));
                r7[c7].Border(Right).Add(NormalParagraph.Alignment(HorizontalAlign.Right)
                        .Add("Офис (квартира)"));
                r7[c8].Border(Bottom | Right)
                    .Add(NormalParagraph.Alignment(HorizontalAlign.Center).Add(""));
            }
            var rowHeight = Cm(0.4);
            {
                var table = section.AddTable().Margin(Top, 4);
                var c1 = table.AddColumn();
                var cellCount = 15;
                for (var i = 0; i < cellCount; i++)
                    table.AddColumn(cellWidth);
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center).Border(Right).Add(NormalParagraph
                    .Add("1.3. Основной государственный регистрационный номер", TimesNewRoman9_5));
                for (var i = 0; i < cellCount; i++)
                    r1[c1.Index + 1 + i].VerticalAlign(VerticalAlign.Center).Border(Top | Right | Bottom)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("111111111111111".CellSubstring(i, cellCount), TimesNewRoman9_5));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                var cellCount = 10;
                for (var i = 0; i < cellCount; i++)
                    table.AddColumn(cellWidth);
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center).Border(Right).Add(NormalParagraph
                    .Add("1.4. Дата внесения записи в государственный реестр", TimesNewRoman9_5));
                for (var i = 0; i < cellCount; i++)
                    r1[c1.Index + 1 + i].VerticalAlign(VerticalAlign.Center).Border(Right)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("3333333333".CellSubstring(i, cellCount), TimesNewRoman9_5));
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                var cellCount = 22;
                for (var i = 0; i < cellCount; i++)
                    table.AddColumn(cellWidth);
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow().Height(rowHeight);
                r1[c1].VerticalAlign(VerticalAlign.Center).Border(Right).Add(NormalParagraph
                    .Add("1.5. ИНН/КПП", TimesNewRoman9_5));
                for (var i = 0; i < cellCount; i++)
                    r1[c1.Index + 1 + i].VerticalAlign(VerticalAlign.Center).Border(Top | Right | Bottom)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("1234123412341234123412".CellSubstring(i, cellCount), TimesNewRoman9_5));
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                .Add("2. Реквизиты нерезидента (нерезидентов)", TimesNewRoman11_5Bold));
            {
	            var table = section.AddTable().Border(BorderWidth).ContentAlign(HorizontalAlign.Center).Font(TimesNewRoman9_5);
                var c1 = table.AddColumn();
                var c2 = table.AddColumn(Px(550));
                var c3 = table.AddColumn(Px(205));
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth;
                var r1 = table.AddRow();
                r1[c1].Rowspan(2).Add(NormalParagraph
                    .Add("Наименование"));
	            r1[c2].Colspan(c3).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
		            .Add("Страна"));
                var r2 = table.AddRow();
                r2[c2].Add(NormalParagraph
                    .Add("наименование"));
                r2[c3].Add(NormalParagraph
                    .Add("код"));
                var r3 = table.AddRow();
                for (var i = 0; i < table.Columns.Count; i++)
                    r3[table.Columns[i]].Add(NormalParagraph
                        .Add($"{i + 1}", TimesNewRoman8));
                for (var i = 0; i < 2; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(NormalParagraph.Add(""));
                    r[c2].Add(NormalParagraph.Add(""));
                    r[c3].Add(NormalParagraph.Add(""));
                }
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                .Add("3. Общие сведения о контракте", TimesNewRoman11_5Bold));
            {
                var table = section.AddTable().Border(BorderWidth).ContentAlign(HorizontalAlign.Center).Font(TimesNewRoman9_5);
                var c1 = table.AddColumn(Px(311));
                var c2 = table.AddColumn(Px(240));
                var c3 = table.AddColumn(Px(235));
                var c4 = table.AddColumn(Px(150));
                var c5 = table.AddColumn(Px(220));
                var c6 = table.AddColumn(Px(455));
                var r1 = table.AddRow();
	            r1[c1].Rowspan(2).Add(NormalParagraph
                    .Add("№"));
                r1[c2].Rowspan(2).Add(NormalParagraph
                    .Add("Дата"));
                r1[c3].Colspan(c4).Add(NormalParagraph
                    .Add("Валюта контракта"));
                r1[c5].Rowspan(2).Add(NormalParagraph
                    .Add("Сумма контракта"));
                r1[c6].Rowspan(2).Add(NormalParagraph
                    .Add("Дата завершения исполнения обязательств по контракту"));
                var r2 = table.AddRow();
                r2[c3].Add(NormalParagraph
                    .Add("наименование"));
                r2[c4].Add(NormalParagraph
                    .Add("код"));
                new[] {c1, c5}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                var r3 = table.AddRow();
                for (var i = 0; i < table.Columns.Count; i++)
                    r3[table.Columns[i]].Add(NormalParagraph
                        .Add($"{i + 1}", TimesNewRoman8));
                for (var i = 0; i < 1; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(NormalParagraph.Add(""));
                    r[c2].Add(NormalParagraph.Add(""));
                    r[c3].Add(NormalParagraph.Add(""));
                    r[c4].Add(NormalParagraph.Add(""));
                    r[c5].Add(NormalParagraph.Add(""));
                    r[c6].Add(NormalParagraph.Add(""));
                }
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                .Add("4. Сведения об оформлении, переводе и закрытии паспорта сделки", TimesNewRoman11_5Bold));
            {
                var table = section.AddTable().Border(BorderWidth).ContentAlign(HorizontalAlign.Center).Font(TimesNewRoman9_5);
                var c1 = table.AddColumn(Px(91));
                var c2 = table.AddColumn(Px(460));
                var c3 = table.AddColumn(Px(380));
                var c4 = table.AddColumn(Px(330));
                var c5 = table.AddColumn(Px(350));
                new[] {c5}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                var r1 = table.AddRow();
	            r1[c1].Add(NormalParagraph
                    .Add("№ п/п"));
                r1[c2].Add(NormalParagraph
                    .Add(@"Регистрационный номер
банка ПС"));
                r1[c3].Add(NormalParagraph
                    .Add(@"Дата принятия паспорта сделки при его переводе"));
                r1[c4].Add(NormalParagraph
                    .Add(@"Дата закрытия паспорта сделки"));
                r1[c5].Add(NormalParagraph
                    .Add(@"Основание закрытия паспорта сделки"));
                var r2 = table.AddRow();
                for (var i = 0; i < table.Columns.Count; i++)
                    r2[table.Columns[i]].Add(NormalParagraph
                        .Add($"{i + 1}", TimesNewRoman8));
                for (var i = 0; i < 2; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(NormalParagraph.Add(""));
                    r[c2].Add(NormalParagraph.Add(""));
                    r[c3].Add(NormalParagraph.Add(""));
                    r[c4].Add(NormalParagraph.Add(""));
                    r[c5].Add(NormalParagraph.Add(""));
                }
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                .Add("5. Сведения о переоформлении паспорта сделки", TimesNewRoman11_5Bold));
            {
                var table = section.AddTable().Border(BorderWidth).ContentAlign(HorizontalAlign.Center).Font(TimesNewRoman9_5);
                var c1 = table.AddColumn(Px(119));
                var c2 = table.AddColumn(Px(323));
                var c3 = table.AddColumn(Px(764));
                var c4 = table.AddColumn(Px(405));
                new[] {c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                var r1 = table.AddRow();
	            r1[c1].Rowspan(2).Add(NormalParagraph
                    .Add(@"№
п/п"));
                r1[c2].Rowspan(2).Add(NormalParagraph
                    .Add("Дата"));
                r1[c3].Colspan(c4).Add(NormalParagraph
                    .Add("Документ, на основании которого внесены изменения в паспорт сделки"));
                var r2 = table.AddRow();
                r2[c3].Add(NormalParagraph
                    .Add("№"));
                r2[c4].Add(NormalParagraph
                    .Add("дата"));
                for (var i = 0; i < 2; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(NormalParagraph.Add(""));
                    r[c2].Add(NormalParagraph.Add(""));
                    r[c3].Add(NormalParagraph.Add(""));
                    r[c4].Add(NormalParagraph.Add(""));
                }
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing)
                .Add("6. Сведения о ранее оформленном", TimesNewRoman11_5Bold));
            {
                var cellCount = 22;
                Table table2;
                {
                    table2 = new Table();
                    for (var i = 0; i < cellCount; i++)
                        table2.AddColumn(cellWidth);
                    var r1 = table2.AddRow().Height(rowHeight);
                    for (var i = 0; i < cellCount; i++)
                        r1[i].VerticalAlign(VerticalAlign.Center).Border(Top | Right | Bottom)
                            .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                                .Add("12345678/1234/1234/1/1".CellSubstring(i, cellCount), TimesNewRoman9_5));
                    r1[0].Border(Left);
                }
                {
                    var table = section.AddTable();
                    var c1 = table.AddColumn();
                    var c2 = table.AddColumn(cellWidth * cellCount + BorderWidth);
                    c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                    var r1 = table.AddRow();
                    r1[c1].Add(NormalParagraph.Margin(Left, Px(41))
                        .Add("паспорте сделки по контракту", TimesNewRoman11_5Bold));
                    r1[c2].VerticalAlign(VerticalAlign.Center).Add(table2);
                }
            }
            section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, smallVerticalSpacing)
                .Add("7. Справочная информация", TimesNewRoman11_5Bold));
            section.Add(new Paragraph()
                .Add("7.1. Способ и дата представления резидентом документов для оформления", TimesNewRoman9_5));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                var c2 = table.AddColumn(Px(90));
                table.AddColumn(Px(40));
                var c4 = table.AddColumn(Px(290));
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add("(переоформления, принятия на обслуживание, закрытия) паспорта сделки", TimesNewRoman9_5));
                r1[c2].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                    .Add("2", TimesNewRoman9_5));
                r1[c4].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                    .Add("2", TimesNewRoman9_5));
            }
            section.Add(new Paragraph().Margin(Top, verticalSpacing)
                .Add("7.2. Способ и дата направления резиденту оформленного (переоформленного,", TimesNewRoman9_5));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                var c2 = table.AddColumn(Px(90));
                table.AddColumn(Px(40));
                var c4 = table.AddColumn(Px(290));
                c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth;
                var r1 = table.AddRow();
                r1[c1].Add(new Paragraph()
                    .Add("принятого на обслуживание, закрытого) паспорта сделки", TimesNewRoman9_5));
                r1[c2].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                    .Add("2", TimesNewRoman9_5));
                r1[c4].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                    .Add("2", TimesNewRoman9_5));
            }
        }

        public static Paragraph NormalParagraph => new Paragraph().Margin(Left | Right, Cm(0.05));

        public static string CellSubstring(this string s, int i, int cellCount)
        {
            if (s == null) return "";
            if (i >= s.Length) return "";
            if (i == cellCount - 1) return s.Substring(i);
            return s.Substring(i, 1);
        }

        public static void Distribute(this Column[] columns, double width)
        {
            var d = width / columns.Length;
            double sum = 0;
            for (var i = 0; i < columns.Length - 1; i++)
            {
                columns[i].Width += d;
                sum += d;
            }
            columns[columns.Length - 1].Width += width - sum;
        }
    }
}