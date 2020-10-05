using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using SharpLayout;

namespace Examples
{
    public static class BandHelper
    {
        public static double Px(double value) => XUnit.FromInch(value / 300d);

        public static Table Band(int columnCount, Indexer indexer, [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            var table = new Table().Font(Styles.Courier15)
                .ContentAlign(HorizontalAlign.Center).ContentVerticalAlign(VerticalAlign.Center);
            for (var i = 0; i < columnCount; i++)
                table.AddColumn(CellWidth);
            var r1 = table.AddRow().Height(CellHeight);
            for (var i = 0; i < columnCount; i++)
            {
                r1[i].Add(new Paragraph()
                    .Add(indexer.Current, line, filePath));
                indexer.Next();
            }
            return table;
        }

        public static double CellHeight => Px(72);
        public static double CellWidth => Px(2333 / 40d);

        public static void VerticalSpacing(this Section section, double value)
        {
            var table = section.AddTable();
            table.AddColumn();
            table.AddRow().Height(value);
        }
    }
}