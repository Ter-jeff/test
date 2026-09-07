using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.IgxlBase;

namespace TestPlanLib.Utility
{
    public static class BinNumberSingletonLegacyHelpers
    {
        public static void ReNameBlockName(ref BinTableRow binTableRow)
        {
            if (!binTableRow.Name.Trim().EndsWithIgnoreCase("HBV") &&
                binTableRow.Name.IndexOf("outsidebincut", StringComparison.OrdinalIgnoreCase) < 0)
            {
                //mode_block_type
                string blockName = "";
                if (IsFuncBlock(binTableRow.Name, ref blockName))
                {
                    binTableRow.Name = binTableRow.Name.Replace(blockName, "Func");
                }

                if (!string.IsNullOrEmpty(binTableRow.ItemList) && IsFuncBlock(binTableRow.ItemList.Split(',').First(), ref blockName))
                {
                    binTableRow.ItemList = binTableRow.ItemList.Replace(blockName, "Func");
                }
            }
        }

        public static bool CompareBinTable(List<BinTableRow> binTableRows, BinTableRow currRow, ref BinTableRow matchRow)
        {
            foreach (BinTableRow findRow in binTableRows)
            {
                string[] findItems = findRow.ItemList.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
                string[] currItems = currRow.ItemList.Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
                if (findItems.Length != currItems.Length)
                {
                    return false;
                }

                for (int i = 0; i < findItems.Length; i++)
                {
                    if (!findItems[i].EqualsIgnoreCase(currItems[i]))
                    {
                        return false;
                    }
                }
                if (!findRow.ItemList.EqualsIgnoreCase(currRow.ItemList))
                {
                    return false;
                }

                var findItemsResult = findRow.Items.Where(x => !string.IsNullOrEmpty(x) && (x.EqualsIgnoreCase("T") || x.EqualsIgnoreCase("F"))).ToList();
                var currItemsResult = currRow.Items.Where(x => !string.IsNullOrEmpty(x) && (x.EqualsIgnoreCase("T") || x.EqualsIgnoreCase("F"))).ToList();
                if (findItemsResult.Count != currItemsResult.Count)
                {
                    return false;
                }

                for (int i = 0; i < findItemsResult.Count; i++)
                {
                    if (!findItemsResult[i].EqualsIgnoreCase(currItemsResult[i]))
                    {
                        return false;
                    }
                }
                if (!findRow.Result.EqualsIgnoreCase(currRow.Result))
                {
                    return false;
                }

                if (!findRow.Op.EqualsIgnoreCase(currRow.Op))
                {
                    return false;
                }

                if (!findRow.Bin.EqualsIgnoreCase(currRow.Bin))
                {
                    return false;
                }

                matchRow = findRow;
                return true;
            }
            return true;
        }

        private static bool IsFuncBlock(string binName, ref string blockName)
        {
            blockName = "";
            if (binName.Trim().EndsWithIgnoreCase("BV") ||
                binName.Trim().EndsWithIgnoreCase("IDS"))
            {
                string[] st = binName.Split('_');
                if (st.Length > 3)
                {
                    blockName = st[^2];
                    return !blockName.EqualsIgnoreCase("TD") && !blockName.EqualsIgnoreCase("Mbist") && !blockName.EqualsIgnoreCase("ELB") && !blockName.EqualsIgnoreCase("ILB");
                }
            }
            return false;
        }
    }
}
