namespace SharpLayout
{
    public class FontFamilyInfo
    {
        public FontSlot Slot { get; }
        public string OriginalName { get; }
        public string Name { get; }

        internal FontFamilyInfo(FontSlot slot, string originalName, string name)
        {
            Slot = slot;
            OriginalName = originalName;
            Name = name;
        }
    }
}