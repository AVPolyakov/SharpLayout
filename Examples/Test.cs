using SharpLayout;
using static Examples.Styles;
using static SharpLayout.Util;

namespace Examples
{
    public static class Test
    {
        public static void AddSection(Document document, TestData data)
        {
            var pageSettings = new PageSettings {
                TopMargin = 0,
                BottomMargin = 0,
                LeftMargin = Cm(1),
                RightMargin = Cm(1)
            };
            var section = document.Add(new Section(pageSettings));
            section.AddHeader(c => {
                switch (c.PageNumber)
                {
                    case 1:
                    {
                        var table = new Table().Font(TimesNewRoman10);
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"First header"));
                        return table;
                    }
                    case 2:
                    {
                        var table = new Table().Font(TimesNewRoman10);
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"Second header
Second header"));
                        return table;
                    }
                    case 3:
                    {
                        var table = new Table().Font(TimesNewRoman10);
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"Header 3
Header 3
Header 3"));
                        return table;
                    }
                    default:
                    {
                        var table = new Table().Font(TimesNewRoman10);
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"Other header
Other header
Other header
Other header"));
                        return table;
                    }
                }
            });
            {
                var table = section.AddTable().Font(TimesNewRoman10).KeepWithNext(true);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 80; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 1, row {i}"));
                }
            }
            {
                var table = section.AddTable().Font(TimesNewRoman10);
                var c1 = table.AddColumn(Cm(5));
                for (var i = 0; i < 160; i++)
                {
                    var r = table.AddRow();
                    r[c1].Add(new Paragraph().Add($"Table 2, row {i}"));
                }
            }
        }
    }
    
    public class TestData
    {
    }
}