using Resources.Fonts;
using SharpLayout;

namespace Resources
{
    public static class FontFamilies
    {
        public static FontFamilyInfo TimesNewRoman => slot.GetFamilyInfo("Times New Roman");
        public static FontFamilyInfo Arial => slot.GetFamilyInfo("Arial");
        public static FontFamilyInfo CourierNew => slot.GetFamilyInfo("Courier New");
        public static FontFamilyInfo Consolas => slot.GetFamilyInfo("Consolas");
        public static FontFamilyInfo Wingdings => slot.GetFamilyInfo("Wingdings");
        public static FontFamilyInfo Wingdings2 => slot.GetFamilyInfo("Wingdings 2");

        private static readonly FontSlot slot = new();
        
        static FontFamilies() => FontResolver.Install(slot, typeof(FontAnchor));
    }
}