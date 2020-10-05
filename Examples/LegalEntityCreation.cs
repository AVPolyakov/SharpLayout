using System.IO;
using Resources;
using SharpLayout;

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
            {
                var indexer = new Indexer(data.FullName);
                for (var i = 0; i < 6; i++)
                    section.Band(71, 775 + 98 * i, 40, indexer);
            }
            {
                var indexer = new Indexer(data.Name);
                for (var i = 0; i < 4; i++)
                    section.Band(70, 1488 + 98 * i, 40, indexer);
            }
            section.Band(579, 2138, 6, data.PostalCode);
            section.Band(1789, 2138, 2, data.Subject);
            section.Band(71, 2388, 10, data.Area);
            {
                var indexer = new Indexer(data.AreaName);
                section.Band(766, 2383, 28, indexer);
                section.Band(71, 2496, 40, indexer);
            }
            section.Band(71, 2676, 10, data.City);
            section.Band(771, 2675, 28, data.CityName);
        }
        
        private static Stream Template()
        {
            var anchorType = typeof(LegalEntityCreationTemplate);
            return anchorType.Assembly.GetManifestResourceStream(anchorType.FullName + ".tiff");
        }
    }
}