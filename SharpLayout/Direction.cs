using System;

namespace SharpLayout
{
    [Flags]
    public enum Direction
    {
        Top = 0x1,
        Bottom = 0x2,
        Left = 0x4,
        Right = 0x8,
        All = Top | Bottom | Left | Right
    }
}