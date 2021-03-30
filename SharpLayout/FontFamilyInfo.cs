namespace SharpLayout
{
    public class FontFamilyInfo
    {
        public FontSlot Slot { get; }
        public string OriginalName { get; }
        public string Name => Slot.Identifier.HasValue ? $"{OriginalName}_{Slot.Identifier}" : OriginalName;

        public FontFamilyInfo(FontSlot slot, string originalName)
        {
            Slot = slot;
            OriginalName = originalName;
        }
    }
}