using SharpLayout;
using static Examples.Styles;
using static SharpLayout.Direction;
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
            section.AddFooter(c => {
                switch (c.PageNumber)
                {
                    case 1:
                    {
                        var table = new Table().Font(TimesNewRoman10)
                            .Margin(Top | Bottom, Px(20));
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"First footer
First footer
First footer
")
                            .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
                        return table;
                    }
                    case 2:
                    {
                        var table = new Table().Font(TimesNewRoman10)
                            .Margin(Top | Bottom, Px(20));
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"Second footer
")
                            .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
                        return table;
                    }
                    default:
                    {
                        var table = new Table().Font(TimesNewRoman10)
                            .Margin(Top | Bottom, Px(20));
                        var c1 = table.AddColumn(Px(500));
                        var r1 = table.AddRow();
                        r1[c1].Add(new Paragraph()
                            .Add(@"Other footer
Other footer
")
                            .Add(rc => $"Page {rc.PageNumber} of {rc.PageCount}"));
                        return table;
                    }
                }
            });
            section.Add(new Paragraph().TextIndent(Cm(1)).Alignment(HorizontalAlign.Justify)
                .Add("Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. " +
                    "Choose composition first when creating new classes from existing classes. Only if " +
                    "inheritance is required by your design should it be used. If you use inheritance where " +
                    "composition will work, your designs will become needlessly complicated. ",
                    TimesNewRoman10));
            {
                var table = section.AddTable().Font(TimesNewRoman10);
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
                for (var i = 0; i < 80; i++)
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