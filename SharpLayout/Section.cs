using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace SharpLayout
{
    public class Section
    {
        public PageSettings PageSettings { get; }
        public List<Table> Tables { get; } = new List<Table>();
        public List<Table> Headers { get; } = new List<Table>();
        public List<Table> Footers { get; } = new List<Table>();

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

        public Table AddHeader([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            Headers.Add(table);
            return table;
        }

        public Table AddFooters([CallerLineNumber] int line = 0)
        {
            var table = new Table(line);
            Footers.Add(table);
            return table;
        }

        public Section Add(Paragraph paragraph, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            var table = new Table(paragraph.KeepWithNext(), line);
            Tables.Add(table);
            var c1 = table.AddColumn(PageSettings.PageWidthWithoutMargins);
            var r1 = table.AddRow();
            r1[c1, line, filePath].Add(paragraph);
            return this;
        }
    }
}