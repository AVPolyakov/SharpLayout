using System.IO;

namespace SharpLayout.Tests
{
    public static class PngImage
    {
        public static void AddSection(Document document)
        {
            var section = document.Add(new Section(new PageSettings()));
            {
                var table = section.AddTable();
                var c1 = table.AddColumn(section.PageSettings.PageWidthWithoutMargins);
                var r1 = table.AddRow();
                r1[c1].Add(new Image()
                    .Content(new VectorImage.VectorImageContent()));
            }
        }

        public class ResourceImageContent : IImageContent
        {
            public Stream CreateStream()
            {
                var anchorType = typeof(Images.Images);
                return anchorType.Assembly
                    .GetManifestResourceStream($"{anchorType.Namespace}.PngImage.png");
            }
        }
    }
}