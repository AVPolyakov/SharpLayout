using System.Collections.Generic;
using System.Runtime.CompilerServices;

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

        public Table AddTable([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            Tables.Add(table);
            return table;
        }
    }
}