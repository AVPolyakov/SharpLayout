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
            var areaName = new Indexer(data.AreaName);
            section
                .Band(71, 775, 40, 6, data.FullName)
                .Band(70, 1488, 40, 4, data.Name)
                .Band(579, 2138, 6, data.PostalCode)
                .Band(1789, 2138, 2, data.Subject)
                .Band(71, 2388, 10, data.Area)
                .Band(766, 2383, 28, areaName)
                .Band(71, 2496, 40, areaName)
                .Band(71, 2676, 10, data.City)
                .Band(771, 2675, 28, data.CityName);
        }
        
        private static Stream Template()
        {
            var anchorType = typeof(LegalEntityCreationTemplate);
            return anchorType.Assembly.GetManifestResourceStream(anchorType.FullName + ".tiff");
        }
    }
}