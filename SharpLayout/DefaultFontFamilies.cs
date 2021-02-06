namespace SharpLayout
{
    public static class DefaultFontFamilies
    {
        public static readonly FontFamilyInfo Roboto = new FontSlot().GetFamilyInfo("Roboto");
        
        static DefaultFontFamilies() => FontResolver.Init();
    }
}