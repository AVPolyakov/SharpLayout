using static SharpLayout.Direction;
using static SharpLayout.Util;
using static SharpLayout.Tests.Styles;
using static SharpLayout.Tests.ContractDealPassport;

namespace SharpLayout.Tests
{
    public static class LoanAgreementDealPassport
    {
        public static void AddSection(Document document)
        {
            var pageSettings = new PageSettings {
                TopMargin = Cm(1),
                BottomMargin = Cm(1),
                LeftMargin = Cm(3),
                RightMargin = Cm(1.5)
            };
            var verticalSpacing = 3d;
            var bigVerticalSpacing = verticalSpacing * 2;
            var smallVerticalSpacing = 2d;
            {
                var section = document.Add(new Section(pageSettings));
                section.Add(new Paragraph().Margin(Bottom, verticalSpacing).Alignment(HorizontalAlign.Right)
                    .Add("Код формы по ОКУД 0406005", TimesNewRoman10));
                section.Add(new Paragraph().Alignment(HorizontalAlign.Right)
                    .Add("Форма 2", TimesNewRoman11));
                section.Add(new Paragraph().Margin(Bottom, verticalSpacing).Alignment(HorizontalAlign.Right)
                    .Add("лист 1", TimesNewRoman10));
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
                    var table = section.AddTable();
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
                    r1[c1].Colspan(c2).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("1.1. Наименование", TimesNewRoman9_5));
                    r1[c3].Colspan(c8).Border(All).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    r2[c1].VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("1.2. Адрес:", TimesNewRoman9_5));
                    r2[c2].Colspan(c4).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Субъект Российской Федерации", TimesNewRoman9_5));
                    r2[c5].Colspan(c8).Border(Bottom | Left | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r3 = table.AddRow();
                    r3[c2].Colspan(c4).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Район", TimesNewRoman9_5));
                    r3[c5].Colspan(c8).Border(Bottom | Left | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r4 = table.AddRow();
                    r4[c2].Colspan(c4).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Город", TimesNewRoman9_5));
                    r4[c5].Colspan(c8).Border(Bottom | Left | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r5 = table.AddRow();
                    r5[c2].Colspan(c4).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Населенный пункт", TimesNewRoman9_5));
                    r5[c5].Colspan(c8).Border(Bottom | Left | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r6 = table.AddRow();
                    r6[c2].Colspan(c4).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Улица (проспект, переулок и т.д.)", TimesNewRoman9_5));
                    r6[c5].Colspan(c8).Border(Bottom | Left | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    var r7 = table.AddRow();
                    r7[c2].Colspan(c3).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("Номер дома (владение)", TimesNewRoman9_5));
                    r7[c4].Border(All).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    r7[c5].Border(Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Right)
                            .Add("Корпус (строение)", TimesNewRoman9_5));
                    r7[c6].Border(Bottom | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    r7[c7].Border(Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Right)
                            .Add("Офис (квартира)", TimesNewRoman9_5));
                    r7[c8].Border(Bottom | Right).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
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
                                .Add("1", TimesNewRoman9_5));
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
                                .Add("3", TimesNewRoman9_5));
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
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn();
                    var c2 = table.AddColumn(Px(550));
                    var c3 = table.AddColumn(Px(205));
                    c1.Width = section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth;
                    var r1 = table.AddRow();
                    r1[c1].Rowspan(2).Add(NormalParagraph
                        .Add("Наименование", TimesNewRoman9_5));
                    r1[c2].Colspan(c3)
                        .Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("Страна", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    r2[c2].Add(NormalParagraph
                        .Add("наименование", TimesNewRoman9_5));
                    r2[c3].Add(NormalParagraph
                        .Add("код", TimesNewRoman9_5));
                    var r3 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r3[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2, r3})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                    .Add("3. Сведения о кредитном договоре", TimesNewRoman11_5Bold));
                section.Add(new Paragraph().Margin(Bottom, smallVerticalSpacing)
                    .Add("3.1. Общие сведения о кредитном договоре", TimesNewRoman10Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Cm(1.4));
                    var c2 = table.AddColumn(Px(140 + 10 + 5));
                    var c3 = table.AddColumn(Px(225));
                    var c4 = table.AddColumn(Px(100));
                    var c5 = table.AddColumn(Cm(1.77));
                    var c6 = table.AddColumn(Cm(2.17));
                    var c7 = table.AddColumn(Px(175));
                    var c8 = table.AddColumn(Px(176));
                    var c9 = table.AddColumn(Px(256 - 5));
                    var r1 = table.AddRow();
                    new[] {c5}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    r1[c1].Rowspan(2).Add(NormalParagraph
                        .Add("№", TimesNewRoman9));
                    r1[c2].Rowspan(2).Add(NormalParagraph
                        .Add("Дата", TimesNewRoman9));
                    r1[c3].Colspan(c4).Add(NormalParagraph
                        .Add(@"Валюта
кредитного договора", TimesNewRoman9));
                    r1[c5].Rowspan(2).Add(NormalParagraph
                        .Add("Сумма кредитного договора", TimesNewRoman9));
                    r1[c6].Rowspan(2).Add(NormalParagraph
                        .Add("Дата завершения исполнения обязательств по кредитному договору", TimesNewRoman9));
                    r1[c7].Colspan(c8).Add(NormalParagraph
                        .Add(@"Особые условия", TimesNewRoman9));
                    r1[c9].Rowspan(2).Add(NormalParagraph
                        .Add("Код срока привлечения (предоставления)", TimesNewRoman9));
                    var r2 = table.AddRow();
                    r2[c3].Add(NormalParagraph
                        .Add("наименование", TimesNewRoman9));
                    r2[c4].Add(NormalParagraph
                        .Add("код", TimesNewRoman9));
                    r2[c7].Add(NormalParagraph
                        .Add("зачисление на счета за\u00A0рубежом", TimesNewRoman9));
                    r2[c8].Add(NormalParagraph
                        .Add("погашение за счет валютной выручки", TimesNewRoman9));
                    var r3 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r3[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2, r3})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    var r4 = table.AddRow();
                    r4[c1].Add(NormalParagraph.Add("", TimesNewRoman9));
                    r4[c2].Add(NormalParagraph.Add("88.88.8888", TimesNewRoman9));
                    r4[c3].Add(NormalParagraph.Add("Российский рубль", TimesNewRoman9));
                    r4[c4].Add(NormalParagraph.Add("USD", TimesNewRoman9));
                    //повод для увеличения правого или левого поля
                    r4[c5].Add(NormalParagraph.Add("1 000 000,00", TimesNewRoman9));
                    r4[c6].Add(NormalParagraph.Add("", TimesNewRoman9));
                    foreach (var column in table.Columns)
                    foreach (var paragraph in r4[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, smallVerticalSpacing)
                    .Add("3.2. Сведения о сумме и сроках привлечения (предоставления) траншей по кредитному договору", TimesNewRoman10Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Px(356));
                    var c2 = table.AddColumn(Px(100 + 20 + 74));
                    var c3 = table.AddColumn(Cm(1.77));
                    var c4 = table.AddColumn(Px(332));
                    var c5 = table.AddColumn(Px(318));
                    var r1 = table.AddRow();
                    new[] {c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    r1[c1].Colspan(c2).Add(NormalParagraph
                        .Add("Валюта кредитного договора", TimesNewRoman9));
                    r1[c3].Rowspan(2).Add(NormalParagraph
                        .Add("Сумма транша", TimesNewRoman9));
                    r1[c4].Rowspan(2).Add(NormalParagraph
                        .Add("Код срока привлечения (предоставления) транша", TimesNewRoman9));
                    r1[c5].Rowspan(2).Add(NormalParagraph
                        .Add(@"Ожидаемая дата поступления
транша", TimesNewRoman9));
                    var r2 = table.AddRow();
                    r2[c1].Add(NormalParagraph
                        .Add("наименование", TimesNewRoman9));
                    r2[c2].Add(NormalParagraph
                        .Add("код", TimesNewRoman9));
                    var r3 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r3[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2, r3})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("Российский рубль", TimesNewRoman9));
                        r[c2].Add(NormalParagraph.Add("USD", TimesNewRoman9));
                        r[c3].Add(NormalParagraph.Add("700000000.00", TimesNewRoman9));
                        r[c4].Add(NormalParagraph.Add("", TimesNewRoman9));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                    .Add("4. Сведения об оформлении, переводе и закрытии паспорта сделки", TimesNewRoman11_5Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Px(91));
                    var c2 = table.AddColumn(Px(460));
                    var c3 = table.AddColumn(Px(380));
                    var c4 = table.AddColumn(Px(330));
                    var c5 = table.AddColumn(Px(350));
                    new[] {c5}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    var r1 = table.AddRow();
                    r1[c1].Add(NormalParagraph
                        .Add("№ п/п", TimesNewRoman9_5));
                    r1[c2].Add(NormalParagraph
                        .Add(@"Регистрационный номер
банка ПС", TimesNewRoman9_5));
                    r1[c3].Add(NormalParagraph
                        .Add(@"Дата принятия паспорта сделки при его переводе", TimesNewRoman9_5));
                    r1[c4].Add(NormalParagraph
                        .Add(@"Дата закрытия паспорта сделки", TimesNewRoman9_5));
                    r1[c5].Add(NormalParagraph
                        .Add(@"Основание закрытия паспорта сделки", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r2[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c4].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c5].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                    .Add("5. Сведения о переоформлении паспорта сделки", TimesNewRoman11_5Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Px(119));
                    var c2 = table.AddColumn(Px(323));
                    var c3 = table.AddColumn(Px(764));
                    var c4 = table.AddColumn(Px(405));
                    new[] {c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    var r1 = table.AddRow();
                    r1[c1].Rowspan(2).Add(NormalParagraph
                        .Add(@"№
п/п", TimesNewRoman9_5));
                    r1[c2].Rowspan(2).Add(NormalParagraph
                        .Add("Дата", TimesNewRoman9_5));
                    r1[c3].Colspan(c4).Add(NormalParagraph
                        .Add("Документ, на основании которого внесены изменения в паспорт сделки", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    r2[c3].Add(NormalParagraph
                        .Add("№", TimesNewRoman9_5));
                    r2[c4].Add(NormalParagraph
                        .Add("дата", TimesNewRoman9_5));
                    foreach (var row in new[] {r1, r2})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c4].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing)
                    .Add("6. Сведения о ранее оформленном", TimesNewRoman10_5Bold));
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
                        //уменьшение шрифта -- повод для увеличения правого или левого поля
                        r1[c1].Add(NormalParagraph.Margin(Left, Px(41))
                            .Add("паспорте сделки по кредитному договору", TimesNewRoman10_5Bold));
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
            {
                var section = document.Add(new Section(pageSettings));
                section.Add(new Paragraph().Margin(Bottom, verticalSpacing).Alignment(HorizontalAlign.Right)
                    .Add("лист 2", TimesNewRoman10));
                section.Add(new Paragraph().Margin(Bottom, verticalSpacing)
                    .Add("8. Специальные сведения о кредитном договоре", TimesNewRoman11_5Bold));
                var indent = Px(61);
                section.Add(new Paragraph().Margin(Bottom, smallVerticalSpacing)
                    .Margin(Left, indent).TextIndent(-indent)
                    .Add(@"8.1. Процентные платежи, предусмотренные кредитным договором
(за исключением платежей по возврату основного долга)", TimesNewRoman10Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Cm(4.3));
                    var c2 = table.AddColumn(Cm(2.25));
                    var c3 = table.AddColumn(Cm(3.25));
                    var c4 = table.AddColumn(Cm(6.3));
                    new[] {c1, c2, c3, c4}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    var r1 = table.AddRow();
                    r1[c1].Add(NormalParagraph
                        .Add(@"Фиксированный размер процентной ставки,
% годовых", TimesNewRoman9_5));
                    r1[c2].Add(NormalParagraph
                        .Add(@"Код ставки ЛИБОР", TimesNewRoman9_5));
                    r1[c3].Add(NormalParagraph
                        .Add(@"Другие методы определения процентной ставки", TimesNewRoman9_5));
                    r1[c4].Add(NormalParagraph
                        .Add(@"Размер процентной надбавки (дополнительных платежей)
к базовой процентной ставке, % годовых", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r2[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c4].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, smallVerticalSpacing)
                    .Margin(Left, indent).TextIndent(-indent)
                    .Add(@"8.2. Иные платежи, предусмотренные кредитным договором (за исключением платежей по возврату
основного долга и процентных платежей, указанных в пункте 8.1)", TimesNewRoman10Bold));
                {
                    var table = section.AddTable();
                    var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                    var r1 = table.AddRow().Height(Cm(0.85));
                    r1[c1].Border(All).VerticalAlign(VerticalAlign.Center)
                        .Add(NormalParagraph.Add("", TimesNewRoman9_5));
                }
                {
                    var table = section.AddTable();
                    table.AddColumn(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                    table.AddRow().Height(verticalSpacing + bigVerticalSpacing);
                }
                {
                    var table = section.AddTable();
                    var c1 = table.AddColumn(Cm(8.3));
                    var c2 = table.AddColumn(Cm(3.9));
                    var c3 = table.AddColumn(Cm(3.9));
                    new[] {c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                    var r1 = table.AddRow();
                    r1[c1].Rowspan(3).Add(new Paragraph()
                        .Margin(Left, indent).TextIndent(-indent)
                        .Add(@"8.3. Сумма задолженности по основному долгу
на дату, предшествующую дате оформления
паспорта сделки", TimesNewRoman10Bold));
                    r1[c2].Border(All).Add(NormalParagraph
                        .Add(@"Код валюты
кредитного договора", TimesNewRoman9_5));
                    r1[c3].Border(Top | Right | Bottom).Add(NormalParagraph
                        .Add("Сумма", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    for (var i = 1; i < table.Columns.Count; i++)
                        r2[table.Columns[i]].Border(Right | Bottom).Add(NormalParagraph
                            .Add($"{i}", TimesNewRoman8));
                    r2[c2].Border(Left);
                    foreach (var row in new[] {r1, r2})
                        foreach (var column in new[] { c2, c3 })
                            foreach (var paragraph in row[column].Paragraphs)
                                paragraph.Alignment(HorizontalAlign.Center);
                    var r3 = table.AddRow();
                    r3[c2].Add(NormalParagraph.Add("USD", TimesNewRoman9_5));
                    r3[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                    foreach (var column in new[] {c2, c3})
                        foreach (var paragraph in r3[column].Border(Right | Bottom).Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    r3[c2].Border(Left);
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, verticalSpacing)
                    .Add("9. Справочная информация о кредитном договоре", TimesNewRoman11_5Bold));
                section.Add(new Paragraph().Margin(Bottom, smallVerticalSpacing)
                    .Add("9.1. Основания заполнения пункта 9.2", TimesNewRoman10Bold));
                {
                    var table = section.AddTable();
                    var c1 = table.AddColumn(Px(176+ 53- 78));
                    var c2 = table.AddColumn(Px(591- 105+ 190));
                    var c3 = table.AddColumn(Px(160));
                    var r1 = table.AddRow();
                    r1[c1].Add(new Paragraph().Margin(Left, Px(25))
                        .Add("9.1.1.", TimesNewRoman9_5));
                    r1[c2].Add(new Paragraph()
                        .Add("Сведения из кредитного договора", TimesNewRoman9_5));
                    r1[c3].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                        .Add("", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    r2[c1].Add(new Paragraph().Margin(Left, Px(25))
                        .Add("9.1.2", TimesNewRoman9_5));
                    r2[c2].Add(new Paragraph()
                        .Add("Оценочные данные", TimesNewRoman9_5));
                    r2[c3].Border(Left | Right | Bottom).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                        .Add("", TimesNewRoman9_5));
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, smallVerticalSpacing)
                    .Add("9.2. Описание графика платежей по возврату основного долга и процентных платежей", TimesNewRoman10Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Px(81));
                    var c2 = table.AddColumn(Px(185));
                    var c3 = table.AddColumn(Px(200));
                    var c4 = table.AddColumn(Px(287.5));
                    var c5 = table.AddColumn(Px(200));
                    var c6 = table.AddColumn(Px(287.5));
                    var c7 = table.AddColumn(Px(330));
                    new[] {c7}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    var r1 = table.AddRow();
                    r1[c1].Rowspan(3).Add(NormalParagraph
                        .Add(@"№
п/п", TimesNewRoman9_5));
                    r1[c2].Rowspan(3).Add(NormalParagraph
                        .Add(@"Код
валюты кредитного договора", TimesNewRoman9_5));
                    r1[c3].Colspan(c6).Add(NormalParagraph
                        .Add(@"Суммы платежей по датам их осуществления,
в единицах валюты кредитного договора", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    r2[c3].Colspan(c4).Add(NormalParagraph
                        .Add("по погашению основного долга", TimesNewRoman9_5));
                    r2[c5].Colspan(c6).Add(NormalParagraph
                        .Add("в счет процентных платежей", TimesNewRoman9_5));
                    r1[c7].Rowspan(3).VerticalAlign(VerticalAlign.Center).Add(NormalParagraph
                            .Add(@"Описание
особых условий", TimesNewRoman9_5));
                    var r3 = table.AddRow();
                    r3[c3].Add(NormalParagraph
                        .Add("дата", TimesNewRoman9_5));
                    r3[c4].Add(NormalParagraph
                        .Add("сумма", TimesNewRoman9_5));
                    r3[c5].Add(NormalParagraph
                        .Add("дата", TimesNewRoman9_5));
                    r3[c6].Add(NormalParagraph
                        .Add("сумма", TimesNewRoman9_5));
                    foreach (var row in new[] {r1, r2, r3})
                        foreach (var column in table.Columns)
                            foreach (var paragraph in row[column].Paragraphs)
                                paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("88.88.8888", TimesNewRoman9_5));
                        r[c4].Add(NormalParagraph.Add("1000000000000.00", TimesNewRoman9_5));
                        r[c5].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c6].Add(NormalParagraph.Add("1000000000000.00", TimesNewRoman9_5));
                        r[c7].Add(NormalParagraph.Add("Тест", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                            foreach (var paragraph in r[column].Paragraphs)
                                paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
                {
                    var table = section.AddTable();
                    table.AddColumn(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                    table.AddRow().Height(verticalSpacing + bigVerticalSpacing);
                }
                {
                    var c1Width = Px(1181);
                    var c2Width = Px(120);
                    var c3Width = Px(305);
                    {
                        var table = section.AddTable();
                        var c1 = table.AddColumn(c1Width);
                        var c2 = table.AddColumn(c2Width);
                        var r1 = table.AddRow();
                        var c3 = table.AddColumn(c3Width);
                        new[] {c1, c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                        r1[c1].Add(new Paragraph()
                            .Add("9.3. Отметка о наличии отношений прямого инвестирования", TimesNewRoman10Bold));
                        r1[c2].Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("", TimesNewRoman9_5));
                    }
                    {
                        var table = section.AddTable();
                        table.AddColumn(section.PageSettings.PageWidthWithoutMargins - BorderWidth);
                        table.AddRow().Height(verticalSpacing + bigVerticalSpacing);
                    }
                    {
                        var table = section.AddTable();
                        var c1 = table.AddColumn(c1Width);
                        var c2 = table.AddColumn(c2Width);
                        var c3 = table.AddColumn(c3Width);
                        new[] {c1, c3}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth);
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add("9.4. Сумма залогового или другого обеспечения", TimesNewRoman10Bold));
                        r1[c2].Colspan(c3).Border(All).Add(NormalParagraph.Alignment(HorizontalAlign.Center)
                            .Add("", TimesNewRoman9_5));
                    }
                }
                section.Add(new Paragraph().Margin(Top, bigVerticalSpacing).Margin(Bottom, smallVerticalSpacing)
                    .Margin(Left, indent).TextIndent(-indent)
                    .Add("9.5. Информация о привлечении резидентом кредита (займа), предоставленного нерезидентами\r\nна\u00A0синдицированной (консорциональной) основе", TimesNewRoman10Bold));
                {
                    var table = section.AddTable().Border(BorderWidth);
                    var c1 = table.AddColumn(Cm(0.8));
                    var c2 = table.AddColumn(Cm(4.5));
                    var c3 = table.AddColumn(Cm(3.45));
                    var c4 = table.AddColumn(Cm(4.05));
                    var c5 = table.AddColumn(Cm(3.3));
                    new[] {c2, c4}.Distribute(section.PageSettings.PageWidthWithoutMargins - table.ColumnsWidth - BorderWidth);
                    var r1 = table.AddRow();
                    r1[c1].Add(NormalParagraph
                        .Add(@"№
п/п", TimesNewRoman9_5));
                    r1[c2].Add(NormalParagraph
                        .Add("Наименование нерезидента", TimesNewRoman9_5));
                    r1[c3].Add(NormalParagraph
                        .Add(@"Код страны
места
нахождения нерезидента", TimesNewRoman9_5));
                    r1[c4].Add(NormalParagraph
                        .Add(@"Предоставляемая сумма денежных средств,
в единицах валюты кредитного договора", TimesNewRoman9_5));
                    r1[c5].Add(NormalParagraph
                        .Add(@"Доля в общей
сумме кредита (займа), %", TimesNewRoman9_5));
                    var r2 = table.AddRow();
                    for (var i = 0; i < table.Columns.Count; i++)
                        r2[table.Columns[i]].Add(NormalParagraph
                            .Add($"{i + 1}", TimesNewRoman8));
                    foreach (var row in new[] {r1, r2})
                    foreach (var column in table.Columns)
                    foreach (var paragraph in row[column].Paragraphs)
                        paragraph.Alignment(HorizontalAlign.Center);
                    for (var i = 0; i < 2; i++)
                    {
                        var r = table.AddRow();
                        r[c1].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c2].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c3].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        r[c4].Add(NormalParagraph.Add("", TimesNewRoman9_5));
                        foreach (var column in table.Columns)
                        foreach (var paragraph in r[column].Paragraphs)
                            paragraph.Alignment(HorizontalAlign.Center);
                    }
                }
            }
        }
    }
}