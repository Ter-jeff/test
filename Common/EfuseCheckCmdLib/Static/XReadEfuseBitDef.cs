using System.Collections.Generic;

using CommonLib.Extension;

using EfuseCheckCmdLib.AlgorithmCheck;
using EfuseCheckCmdLib.EFuse.EFuseApp;

namespace EfuseCheckCmdLib.Static
{
    public class XReadEfuseBitDef
    {
        public static List<XBitDefRow> BitDefTable { get; set; } = [];

        public XReadEfuseBitDef()
        {
            BitDefTable.Clear();
        }

        public static bool Contains(string name)
        {
            foreach (XBitDefRow oneRow in BitDefTable)
            {
                if (oneRow.Name == name)
                {
                    return true;
                }
            }

            return false;
        }

        public static XBitDefRow? GetRow(string name, string block)
        {
            string tmpName = name.Replace("(", @"\(").Replace(")", @"\)").ToLower();
            foreach (XBitDefRow oneRow in BitDefTable)
            {
                if (oneRow.Name.Split('#')[1].EqualsIgnoreCase(tmpName) && oneRow.Block == block)
                {
                    return oneRow;
                }
            }
            foreach (XBitDefRow oneRow in BitDefTable)
            {
                if (oneRow.Name.Split('#')[1].EqualsIgnoreCase(tmpName))
                {
                    return oneRow;
                }
            }

            return null;
        }

        public static EnumValueType GetType(string name, string block)
        {
            string tmpName = name.Replace("(", @"\(").Replace(")", @"\)");
            foreach (XBitDefRow oneRow in BitDefTable)
            {
                if (oneRow.Name.Split('#')[1].EqualsIgnoreCase(tmpName) && oneRow.Block == block)
                {
                    return oneRow.Type;
                }
            }

            return EnumValueType.Other;
        }
    }
}
