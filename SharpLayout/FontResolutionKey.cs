using System;

namespace SharpLayout
{
    internal readonly struct FontResolutionKey : IEquatable<FontResolutionKey>
    {
        private readonly string familyName;
        private readonly bool isBold;
        private readonly bool isItalic;

        public FontResolutionKey(string familyName, bool isBold, bool isItalic)
        {
            this.familyName = familyName;
            this.isBold = isBold;
            this.isItalic = isItalic;
        }

        public bool Equals(FontResolutionKey other)
        {
            return familyName == other.familyName && isBold == other.isBold && isItalic == other.isItalic;
        }

        public override bool Equals(object obj)
        {
            return obj is FontResolutionKey other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = familyName != null ? familyName.GetHashCode() : 0;
                hashCode = (hashCode * 397) ^ isBold.GetHashCode();
                hashCode = (hashCode * 397) ^ isItalic.GetHashCode();
                return hashCode;
            }
        }
    }
}