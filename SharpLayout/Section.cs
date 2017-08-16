using System.Collections.Generic;

namespace SharpLayout
{
    public class Section
    {
        public PageSettings PageSettings { get; }
        public List<Table> Tables { get; } = new List<Table>();

        public Section(PageSettings pageSettings)
        {
            PageSettings = pageSettings;
        }

        public void Add(Table table) => Tables.Add(table);
    }
}