using System;
using PdfSharp.Drawing;
using static SharpLayout.Direction;

namespace SharpLayout
{
    public class Image : IElement
    {
        public T Match<T>(Func<Paragraph, T> paragraph, Func<Table, T> table, Func<Image, T> image) => image(this);

        private double? topMargin;
        public double? TopMargin() => topMargin;
        public Image TopMargin(double? value)
        {
            topMargin = value;
            return this;
        }

        private double? height;
        public double? Height() => height;
        public Image Height(double? value)
        {
            height = value;
            return this;
        }

        private double? width;
        public double? Width() => width;
        public Image Width(double? value)
        {
            width = value;
            return this;
        }

        private double? leftMargin;
        public double? LeftMargin() => leftMargin;
        public Image LeftMargin(double? value)
        {
            leftMargin = value;
            return this;
        }

        private double? rightMargin;
        public double? RightMargin() => rightMargin;
        public Image RightMargin(double? value)
        {
            rightMargin = value;
            return this;
        }

        private double? bottomMargin;
        public double? BottomMargin() => bottomMargin;
        public Image BottomMargin(double? value)
        {
            bottomMargin = value;
            return this;
        }

        public Image Margin(Direction direction, double value)
        {
            if (direction.HasFlag(Left)) LeftMargin(value);
            if (direction.HasFlag(Right)) RightMargin(value);
            if (direction.HasFlag(Top)) TopMargin(value);
            if (direction.HasFlag(Bottom)) BottomMargin(value);
            return this;
        }

        private Option<IImageContent> content;
        public Option<IImageContent> Content() => content;
        public Image Content(IImageContent value)
        {
            content = value.ToOption();
            return this;
        }

        private HorizontalAlign? alignment;
        public HorizontalAlign? Alignment() => alignment;
        public Image Alignment(HorizontalAlign? value)
        {
	        alignment = value;
	        return this;
        }
    }

    public interface IImageContent
    {
        T Process<T>(Func<XImage, T> func);
    }
}