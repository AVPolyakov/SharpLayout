using System;
using System.Collections.Generic;
using System.Linq;
using PdfSharp.Drawing;

namespace SharpLayout
{
    public class Drawer
    {
        private readonly IGraphics graphics;
        private readonly List<Item> items = new List<Item>();

        public Drawer(IGraphics graphics)
        {
            this.graphics = graphics;            
        }

        public void DrawString(string s, Font font, XBrush brush, double x, double y)
        {
            Add(DrawType.Foreground, () => graphics.DrawString(s, font, brush, x, y));
        }

        public void DrawLine(XPen pen, double x1, double y1, double x2, double y2)
        {
            Add(DrawType.Foreground, () => graphics.DrawLine(pen, x1, y1, x2, y2));
        }

        public void DrawImage(Image image, double x, double y)
        {
            Add(DrawType.Background, () => {
                var content = image.Content();
                if (content.HasValue)
                    if (!image.Height().HasValue && !image.Width().HasValue)
                        graphics.DrawImage(content.Value, x, y);
                    else
                        graphics.DrawImage(content.Value, x, y,
                            width => image.Width().GetValueOrDefault(width),
                            height => image.Height().GetValueOrDefault(height));
            });
        }

        public void DrawRectangle(XBrush brush, double x, double y, double width, double height,
            DrawType drawType = DrawType.Foreground)
        {
            Add(drawType, () => graphics.DrawRectangle(brush, x, y, width, height));
        }

        public void Flush()
        {
            Flush(DrawType.Background);
            Flush(DrawType.Foreground);
        }

        private void Flush(DrawType drawType)
        {
            foreach (var item in items.Where(_ => _.DrawType == drawType))
                item.Action();
        }

        private void Add(DrawType drawType, Action action) => items.Add(new Item(drawType, action));

        private class Item
        {
            public readonly DrawType DrawType;
            public readonly Action Action;

            public Item(DrawType drawType, Action action)
            {
                DrawType = drawType;
                Action = action;
            }
        }
    }

    public enum DrawType
    {
        Background,
        Foreground
    }
}