using System;

namespace SharpLayout
{
    public class FontSlot
    {
        public Guid? Identifier;

        public FontSlot(Guid? identifier) => Identifier = identifier;

        public FontSlot() : this(Guid.NewGuid())
        {
        }

        public FontFamilyInfo FamilyInfo(string familyName) => new(this, familyName);
    }
}