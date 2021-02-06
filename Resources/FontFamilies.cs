using Resources.Fonts;
using SharpLayout;

namespace Resources
{
    public static class FontFamilies
    {
        public static FontFamilyInfo TimesNewRoman => slot.FamilyInfo("Times New Roman");
        public static FontFamilyInfo Arial => slot.FamilyInfo("Arial");
        public static FontFamilyInfo CourierNew => slot.FamilyInfo("Courier New");
        public static FontFamilyInfo Consolas => slot.FamilyInfo("Consolas");
        public static FontFamilyInfo Wingdings => slot.FamilyInfo("Wingdings");
        public static FontFamilyInfo Wingdings2 => slot.FamilyInfo("Wingdings 2");

        private static readonly FontSlot slot = new();
        
        static FontFamilies() => FontResolver.Install(slot, typeof(FontAnchor));
    }
}