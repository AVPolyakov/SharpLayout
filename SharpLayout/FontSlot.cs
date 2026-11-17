namespace SharpLayout
{
    public class FontSlot
    {
        public FontFamilyInfo FamilyInfo(string familyName) => new(this, familyName);
    }
}