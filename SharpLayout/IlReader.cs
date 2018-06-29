using System;
using System.Collections.Generic;
using System.Reflection;
using System.Reflection.Emit;

namespace SharpLayout
{
    /// <summary>
    /// http://www.codeproject.com/KB/cs/sdilreader.aspx
    /// </summary>
    internal static class IlReader
    {
        public static IEnumerable<Tuple<OpCode, int?>> Read(MethodBase methodInfo)
        {
            var methodBody = methodInfo.GetMethodBody();
            if (methodBody == null) yield break;
            var ilAsByteArray = methodBody.GetILAsByteArray();
            var position = 0;
            while (position < ilAsByteArray.Length)
            {
                OpCode opCode;
                ushort value = ilAsByteArray[position++];
                if (value == 0xfe)
                {
                    value = ilAsByteArray[position++];
                    opCode = multiByteOpCodes[value];
                }
                else
                    opCode = singleByteOpCodes[value];
                var metadataToken = Read(opCode, ilAsByteArray, ref position);
                yield return Tuple.Create(opCode, metadataToken);
            }
        }

        private static int? Read(OpCode opCode, byte[] ilAsByteArray, ref int position)
        {
            switch (opCode.OperandType)
            {
                case OperandType.InlineBrTarget:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineField:
                    return ReadInt32(ilAsByteArray, ref position);
                case OperandType.InlineMethod:
                    return ReadInt32(ilAsByteArray, ref position);
                case OperandType.InlineSig:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineTok:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineType:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineI:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineI8:
                    ReadInt64(ref position);
                    return new int?();
                case OperandType.InlineNone:
                    return new int?();
                case OperandType.InlineR:
                    ReadDouble(ref position);
                    return new int?();
                case OperandType.InlineString:
                    ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineSwitch:
                    var count = ReadInt32(ilAsByteArray, ref position);
                    for (var i = 0; i < count; i++) ReadInt32(ilAsByteArray, ref position);
                    return new int?();
                case OperandType.InlineVar:
                    ReadUInt16(ref position);
                    return new int?();
                case OperandType.ShortInlineBrTarget:
                    ReadSByte(ref position);
                    return new int?();
                case OperandType.ShortInlineI:
                    ReadSByte(ref position);
                    return new int?();
                case OperandType.ShortInlineR:
                    ReadSingle(ref position);
                    return new int?();
                case OperandType.ShortInlineVar:
                    ReadByte(ref position);
                    return new int?();
                default:
                    throw new InvalidOperationException();
            }
        }

        private static void ReadUInt16(ref int position) => position += 2;
        private static int ReadInt32(byte[] bytes, ref int position) 
            => bytes[position++] | bytes[position++] << 8 | bytes[position++] << 0x10 | bytes[position++] << 0x18;
        private static void ReadInt64(ref int position) => position += 8;
        private static void ReadDouble(ref int position) => position += 8;
        private static void ReadSByte(ref int position) => position++;
        private static void ReadByte(ref int position) => position++;
        private static void ReadSingle(ref int position) => position += 4;

        static IlReader()
        {
            singleByteOpCodes = new OpCode[0x100];
            multiByteOpCodes = new OpCode[0x100];
            foreach (var fieldInfo in typeof (OpCodes).GetFields())
                if (fieldInfo.FieldType == typeof (OpCode))
                {
                    var opCode = (OpCode) fieldInfo.GetValue(null);
                    var value = unchecked((ushort) opCode.Value);
                    if (value < 0x100)
                        singleByteOpCodes[value] = opCode;
                    else
                    {
                        if ((value & 0xff00) != 0xfe00)
                            throw new ApplicationException("Invalid OpCode.");
                        multiByteOpCodes[value & 0xff] = opCode;
                    }
                }
        }

        private static readonly OpCode[] multiByteOpCodes;
        private static readonly OpCode[] singleByteOpCodes;
    }
}