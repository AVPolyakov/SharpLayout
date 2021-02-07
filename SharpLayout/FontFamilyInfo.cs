namespace SharpLayout
{
    public class FontFamilyInfo
    {
        public FontSlot Slot { get; }
        public string OriginalName { get; }
        public string Name => $"{OriginalName}_{Slot.Identifier}";

        internal FontFamilyInfo(FontSlot slot, string originalName)
        {
            Slot = slot;
            OriginalName = originalName;
        }
    }
}