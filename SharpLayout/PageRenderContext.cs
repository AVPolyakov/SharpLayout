namespace SharpLayout
{
    public class PageRenderContext
    {
        private readonly int _pageIndex;

        public PageRenderContext(int pageIndex)
        {
            _pageIndex = pageIndex;
        }

        public int PageNumber => _pageIndex + 1;
    }
}