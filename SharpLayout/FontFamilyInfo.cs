namespace SharpLayout
{
    public class FontFamilyInfo
    {
        public FontSlot Slot { get; }
        public string Name { get; }
        public string FullName => $"{Name}_{Slot.Identifier}";

        internal FontFamilyInfo(FontSlot slot, string name)
        {
            Slot = slot;
            Name = name;
        }
    }
}