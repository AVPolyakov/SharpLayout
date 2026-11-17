namespace SharpLayout
{
    public class FontFamilyInfo
    {
        public FontSlot Slot { get; }
        public string OriginalName { get; }
        public string Name { get; }

        public FontFamilyInfo(FontSlot slot, string originalName)
        {
            Slot = slot;
            OriginalName = originalName;
            Name = originalName;
        }
    }
}