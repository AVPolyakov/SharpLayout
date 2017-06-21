namespace SharpLayout
{
    public class Column
    {
        public double Width { get; set; }

        public int Index { get; }

        internal Column(int index)
        {
            Index = index;
        }
    }
}