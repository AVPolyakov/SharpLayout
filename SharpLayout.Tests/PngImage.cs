namespace SharpLayout.Tests
{
    public static class PngImage
    {
        public static void AddSection(Document document)
        {
            var section = document.Add(new Section(new PageSettings()));
        }
    }
}