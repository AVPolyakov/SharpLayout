namespace SharpLayout
{
    public interface IPageFilter
    {
        void OnNewPage();
        bool PageMustBeAdd { get; }
    }

    public class AllPageFilter: IPageFilter
    {
        public static readonly IPageFilter Instance = new AllPageFilter();
        
        private AllPageFilter()
        {
        }

        public void OnNewPage()
        {
        }

        public bool PageMustBeAdd => true;
    }

    public class OnePageFilter: IPageFilter
    {
        private readonly int pageNumber;
        private int pageIndex;

        public OnePageFilter(int pageNumber)
        {
            this.pageNumber = pageNumber;
        }

        public void OnNewPage() => pageIndex++;

        public bool PageMustBeAdd => pageIndex == pageNumber;
    }
}