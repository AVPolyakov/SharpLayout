using System;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SharpLayout
{
    public struct FontKey : IEquatable<FontKey>
    {
        public string Name { get; }
        public double Size { get; }
        public XFontStyle Style { get; }
        public PdfFontEncoding FontEncoding { get; }

        public FontKey(string name, double size, XFontStyle style, PdfFontEncoding fontEncoding)
        {
            Name = name;
            Size = size;
            Style = style;
            FontEncoding = fontEncoding;
        }

        public bool Equals(FontKey other)
        {
            return string.Equals(Name, other.Name) && Size.Equals(other.Size) && Style == other.Style && FontEncoding == other.FontEncoding;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            return obj is FontKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Name != null ? Name.GetHashCode() : 0;
                hashCode = hashCode * 397 ^ Size.GetHashCode();
                hashCode = hashCode * 397 ^ (int) Style;
                hashCode = hashCode * 397 ^ (int) FontEncoding;
                return hashCode;
            }
        }
    }
}