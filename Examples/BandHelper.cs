using System.Runtime.CompilerServices;
using PdfSharp.Drawing;
using SharpLayout;

namespace Examples
{
    public static class BandHelper
    {
        public static double Px(double value) => XUnit.FromInch(value / 300d);

        public static Section Band(this Section section, double x, double y, int columnCount, int rowCount, Indexer indexer,
            [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            for (var i = 0; i < rowCount; i++)
                section.Band(x, y + 98 * i, columnCount, indexer, line, filePath);
            return section;
        }

        public static Section Band(this Section section, double x, double y, int columnCount, Indexer indexer,
            [CallerLineNumber] int line = 0, [CallerFilePath] string filePath = "")
        {
            var table = section.AddTable().Font(Styles.Courier15)
                .ContentAlign(HorizontalAlign.Center).ContentVerticalAlign(VerticalAlign.Center)
                .Location(Px(x), Px(y));
            for (var i = 0; i < columnCount; i++)
                table.AddColumn(Px(2333 / 40d));
            var r1 = table.AddRow().Height(Px(72));
            for (var i = 0; i < columnCount; i++)
            {
                r1[i].Add(new Paragraph()
                    .Add(indexer.Current, line, filePath));
                indexer.Next();
            }
            return section;
        }
    }
}