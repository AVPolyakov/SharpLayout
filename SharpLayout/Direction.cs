using System;

namespace SharpLayout
{
    [Flags]
    public enum Direction
    {
        Left = 0x1,
        Right = 0x2,
        Top = 0x4,
        Bottom = 0x8,
    }
}