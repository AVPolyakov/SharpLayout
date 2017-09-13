using System.Linq;
using PdfSharp;
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
            var cellMargin = Cm(0.05);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin;
                var r1 = table.AddRow();
                var paragraph = TimesNewRoman10_5("Код формы по ОКУД 0406009");
                paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                paragraph.Alignment = HorizontalAlignment.Right;
                paragraph.BottomMargin = Px(11);
                r1[c1].Add(paragraph);
            }
            var rowHeight = Cm(0.49);
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(6.3));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = rowHeight;
                {
                    var paragraph = TimesNewRoman9_5("Наименование уполномоченного банка");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    var cell = r1[c1];
                    cell.Add(paragraph);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                }
                {
                    var cell = r1[c2];
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                r2.Height = rowHeight;
                {
                    var paragraph = TimesNewRoman9_5("Наименование резидента");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    var cell = r2[c1];
                    cell.Add(paragraph);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                }
                {
                    var cell = r2[c2];
                    cell.LeftBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin;
                var r1 = table.AddRow();
                var paragraph = TimesNewRoman11_5Bold("СПРАВКА О ВАЛЮТНЫХ ОПЕРАЦИЯХ");
                paragraph.Alignment = HorizontalAlignment.Center;
                paragraph.TopMargin = Px(41);
                r1[c1].Add(paragraph);
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1146));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                {
                    var paragraph = TimesNewRoman11_5Bold("от  ");
                    paragraph.TopMargin = Px(30 - 18);
                    r1[c2].Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                table.AddColumn(Px(1199));
                var c2 = table.AddColumn(Px(200));
                var r1 = table.AddRow();
                r1.Height = Px(28);
                r1[c2].TopBorder = BorderWidth;
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Cm(8));
                var c2 = table.AddColumn(Px(405));
                var c3 = table.AddColumn();
                var c4 = table.AddColumn(Px(325));
                c3.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = rowHeight;
                {
                    var paragraph = TimesNewRoman9_5("Номер счета резидента в уполномоченном банке");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    var cell = r1[c1];
                    cell.Add(paragraph);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                }
                {
                    var cell = r1[c2];
                    cell.Colspan(c4);
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
                var r2 = table.AddRow();
                r2.Height = rowHeight;
                {
                    var paragraph = TimesNewRoman9_5("Код страны банка-нерезидента");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    var cell = r2[c1];
                    cell.Add(paragraph);
                    cell.VerticalAlignment = VerticalAlignment.Center;
                }
                {
                    var cell = r2[c2];
                    cell.LeftBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r2[c3];
                    var paragraph = TimesNewRoman9_5("Признак корректировки");
                    paragraph.Alignment = HorizontalAlignment.Right;
                    paragraph.LeftMargin = cellMargin;
                    paragraph.RightMargin = Cm(0.2) + cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r2[c4];
                    cell.LeftBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.Alignment = HorizontalAlignment.Center;
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    cell.Add(paragraph);
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    table.Columns.Sum(_ => _.Width);
                table.AddRow().Height = Px(35);
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
                c2.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    BorderWidth - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                {
                    var cell = r1[c1];
                    cell.Rowspan = 2;
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("№ п/п"));
                }
                {
                    var cell = r1[c2];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Уведомление, распоряжение, расчетный или иной документ"));
                }
                {
                    var cell = r1[c3];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Дата операции"));
                }
                {
                    var cell = r1[c4];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Признак платежа"));
                }
                {
                    var cell = r1[c5];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Код вида валютной операции"));
                }
                {
                    var cell = r1[c6];
                    cell.Colspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Сумма операции"));
                }
                {
                    var cell = r1[c8];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9(@"Номер ПС
или номер и (или) дата договора (контракта)"));
                }
                {
                    var cell = r1[c9];
                    cell.Colspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9(@"Сумма операции
в единицах валюты контракта (кредитного договора)"));
                }
                {
                    var cell = r1[c11];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Срок возврата аванса"));
                }
                {
                    var cell = r1[c12];
                    cell.Rowspan = 2;
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("Ожидаемый срок"));
                }
                var r2 = table.AddRow();
                {
                    var cell = r2[c6];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("код валюты"));
                }
                {
                    var cell = r2[c7];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("сумма"));
                }
                {
                    var cell = r2[c9];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("код валюты"));
                }
                {
                    var cell = r2[c10];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9("сумма"));
                }
                var r3 = table.AddRow();
                r3[c1].LeftBorder = BorderWidth;
                for (var i = 0; i < table.Columns.Count; i++)
                {
                    var cell = r3[table.Columns[i]];
                    cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.Add(TimesNewRoman9($"{i + 1}"));
                }
                foreach (var row in new[] {r1, r2, r3})
                foreach (var column in table.Columns)
                {
                    var cell = row[column];
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                        paragraph.Alignment = HorizontalAlignment.Center;
                    }
                }
                for (var i = 0; i < 3; i++)
                {
                    var row = table.AddRow();
                    row.Height = Cm(0.46);
                    row[c1].LeftBorder = BorderWidth;
                    row[c1].Add(TimesNewRoman9_5($"{i + 1}"));
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
                    {
                        var cell = row[column];
                        cell.RightBorder = cell.BottomBorder = BorderWidth;
                        cell.VerticalAlignment = VerticalAlignment.Center;
                        foreach (var paragraph in cell.Paragraphs)
                        {
                            paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                            paragraph.Alignment = HorizontalAlignment.Center;
                        }
                    }
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(363));
                var r1 = table.AddRow();
                r1.Height = Px(41);
                r1[c1].BottomBorder = BorderWidth;
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin);
                var r1 = table.AddRow();
                var paragraph = TimesNewRoman9_5("Примечание.");
                paragraph.TopMargin = Px(1);
                paragraph.BottomMargin = Px(4);
                r1[c1].Add(paragraph);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(Px(255));
                var c2 = table.AddColumn();
                c2.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    BorderWidth - table.Columns.Sum(_ => _.Width);
                var r1 = table.AddRow();
                r1.Height = Cm(0.46);
                {
                    var cell = r1[c1];
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman9("№ строки");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                {
                    var cell = r1[c2];
                    cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman9("Содержание");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
                for (var i = 0; i < 2; i++)
                {
                    var row = table.AddRow();
                    row.Height = Cm(0.46);
                    row[c1].LeftBorder = BorderWidth;
                    foreach (var column in table.Columns)
                    {
                        var cell = row[column];
                        cell.RightBorder = cell.BottomBorder = BorderWidth;
                        cell.VerticalAlignment = VerticalAlignment.Center;
                    }
                    {
                        var paragraph = TimesNewRoman9_5("");
                        paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                        paragraph.Alignment = HorizontalAlignment.Center;
                        row[c1].Add(paragraph);
                    }
                    {
                        var paragraph = TimesNewRoman9_5("");
                        paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                        paragraph.Alignment = HorizontalAlignment.Center;
                        row[c2].Add(paragraph);
                    }
                }
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin);
                var r1 = table.AddRow();
                var paragraph = TimesNewRoman9_5("Информация уполномоченного банка");
                paragraph.TopMargin = Px(3);
                paragraph.BottomMargin = Px(4);
                r1[c1].Add(paragraph);
            }
            {
                var table = section.AddTable();
                var c1 = table.AddColumn();
                c1.Width = pageSettings.PageWidth - pageSettings.LeftMargin - pageSettings.RightMargin -
                    BorderWidth;
                var r1 = table.AddRow();
                r1.Height = Cm(0.46);
                {
                    var cell = r1[c1];
                    cell.LeftBorder = cell.TopBorder = cell.RightBorder = cell.BottomBorder = BorderWidth;
                    cell.VerticalAlignment = VerticalAlignment.Center;
                    var paragraph = TimesNewRoman9_5("");
                    paragraph.LeftMargin = paragraph.RightMargin = cellMargin;
                    paragraph.Alignment = HorizontalAlignment.Center;
                    cell.Add(paragraph);
                }
            }
        }
    }
}