namespace SharpLayout
{
    public static class DefaultFontFamilies
    {
        public static readonly FontFamilyInfo Roboto = new FontSlot().FamilyInfo("Roboto");
        
        static DefaultFontFamilies() => FontResolver.Init();
    }
}