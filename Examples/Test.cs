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
        }
    }
    
    public class TestData
    {
    }
}