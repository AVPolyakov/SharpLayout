using System.IO;
using Resources;
using SharpLayout;
using static Examples.BandHelper;
using static SharpLayout.Direction;

namespace Examples
{
    public static class LegalEntityCreation
    {
        public static void AddSection(Document document, LegalEntityCreationData data)
        {
            var pageSettings = new PageSettings
            {
                TopMargin = 0,
                BottomMargin = 0,
                LeftMargin = 0,
                RightMargin = 0
            };
            var section = document.Add(new Section(pageSettings))
                .Template(Template);
            var rowSpacing = Px(26);
            section.VerticalSpacing(Px(775));
            {
                var indexer = new Indexer(data.FullName);
                for (var i = 0; i < 6; i++)
                    section.Add(Band(40, indexer)
                        .Margin(Left, Px(71)).Margin(Bottom, rowSpacing));
            }
            section.VerticalSpacing(Px(151) - rowSpacing);
            {
                var indexer = new Indexer(data.Name);
                for (var i = 0; i < 4; i++)
                    section.Add(Band(40, indexer)
                        .Margin(Left, Px(70)).Margin(Bottom, rowSpacing));
            }
            section.VerticalSpacing(Px(284) - rowSpacing);
            {
                var table = section.AddTable()
                    .Margin(Left, Px(579));
                var c1 = table.AddColumn(Px(1210));
                var c2 = table.AddColumn();
                var r1 = table.AddRow();
                r1[c1].Add(Band(6, new Indexer(data.PostalCode)));
                r1[c2].Add(Band(2, new Indexer(data.Subject)));
            }
            section.VerticalSpacing(Px(173));
            var areaIndexer = new Indexer(data.AreaName);
            {
                var table = section.AddTable()
                    .Margin(Left, Px(71));
                var c1 = table.AddColumn(Px(695));
                var c2 = table.AddColumn();
                var r1 = table.AddRow();
                r1[c1].Add(Band(10, new Indexer(data.Area))
                    .Margin(Top, Px(5)));
                r1[c2].Add(Band(28, areaIndexer));
            }
            section.VerticalSpacing(Px(36));
            {
                section.Add(Band(40, areaIndexer)
                    .Margin(Left, Px(71)));
            }
            section.VerticalSpacing(Px(107));
            {
                var table = section.AddTable()
                    .Margin(Left, Px(71));
                var c1 = table.AddColumn(Px(700));
                var c2 = table.AddColumn();
                var r1 = table.AddRow();
                r1[c1].Add(Band(10, new Indexer(data.City))
                    .Margin(Top, Px(1)));
                r1[c2].Add(Band(28, new Indexer(data.CityName)));
            }
        }
        
        private static Stream Template()
        {
            var anchorType = typeof(LegalEntityCreationTemplate);
            return anchorType.Assembly.GetManifestResourceStream(anchorType.FullName + ".tiff");
        }
    }
}