using System;

namespace SharpLayout
{
    public class FontSlot
    {
        public string Identifier = Guid.NewGuid().ToString();

        public FontFamilyInfo GetFamilyInfo(string familyName) => new(this, familyName);
    }
}