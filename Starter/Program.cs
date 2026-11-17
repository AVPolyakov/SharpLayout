using System;
using System.Diagnostics;
using Examples;
using PdfSharp;
using SharpLayout;
using static SharpLayout.Direction;
using static SharpLayout.Util;

namespace Starter
{
    static class Program
    {
        static void Main()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            const int count = 3000;

            for (var j = 1; j <= 5; j++)
                M1(count);
            for (var j = 1; j <= 20; j++)
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();
                M1(count);
                stopwatch.Stop();
                Console.WriteLine($"{stopwatch.ElapsedMilliseconds}");
            }
        }

        private static void M1(int count)
        {
            var document = new Document();
            var section = document.Add(new Section(new PageSettings
            {
                TopMargin = Cm(1.2),
                BottomMargin = Cm(1),
                LeftMargin = Cm(2),
                RightMargin = Cm(1),
                Orientation = PageOrientation.Landscape
            }));
            var table = section.AddTable().Border(Styles.BorderWidth);
            var columnCount = 14;
            for (var i = 0; i < columnCount; i++)
                table.AddColumn();
            table.Columns.ToArray().Distribute(section.PageSettings.PageWidthWithoutMargins);
            for (var j = 0; j < count; j++)
            {
                table.AddRow(r => {
                    for (var i = 0; i < columnCount; i++)
                    {
                        r[table.Columns[i]].Add(new Paragraph().Margin(Left | Right, Cm(0.05))
                            .Add("QWEQWE QWETQWTQWRTWERT QWETQWET QWEQWR ASDASF", Styles.TimesNewRoman8));
                    }
                });
            }

            document.CreatePdf();
        }
    }
}
