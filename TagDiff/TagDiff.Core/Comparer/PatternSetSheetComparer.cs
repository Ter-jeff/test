using System;
using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Common;

using TagDiffCore.Utility;

namespace TagDiff.Core.Comparer
{
    public class PatternSetSheetComparer : IgxlSheetComparerBase
    {
        public Dictionary<int, int> Dic22To23 = new()
        {
              {0,0},{1,1},{2,-1},{3,2},{4,3},{5,4},{6,5},{7,6},{8,7},{9,8},{10,9}
        };

        public PatternSetSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Pattern Set", ["Comment"])
        {
            KeyColHeaders = ["Pattern Set"];
            SheetType = EnumSheetType.DTPatternSetSheet;
            BaseSheetVersion = SheetProvider.GetSheetVersion(BaseSheet);
            if (CompSheet.Count != 0)
            {
                CompareSheetVersion = SheetProvider.GetSheetVersion(CompSheet);
            }
        }

        protected override bool CompareSheetRow(Location baseLocation, Location compLocation, HashSet<int>? skips = null)
        {
            var dic = new Dictionary<int, int>();
            if (BaseSheetVersion == CompareSheetVersion)
            {
                dic = [];
            }
            else if (BaseSheetVersion == "2.2" && CompareSheetVersion == "2.3")
            {
                dic = Dic22To23;
            }
            else if (BaseSheetVersion == "2.3" && CompareSheetVersion == "2.2")
            {
                var reversed = Dic22To23.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);
                dic = reversed;
            }

            bool result = true;
            for (int col = baseLocation.Col; col <= BaseSheetEndCol; col++)
            {
                if (skips != null && skips.Contains(col))
                {
                    continue;
                }

                string baseContext = GetCellSafe(BaseSheet, baseLocation.Row, col);
                string comContext;
                if (dic.Count == 0)
                {
                    comContext = GetCellSafe(CompSheet, compLocation.Row, col);
                }
                else
                {
                    int offset = dic[col];
                    comContext = GetCellSafe(CompSheet, compLocation.Row, offset);
                }
                if (!baseContext.Trim('=').EqualsIgnoreCase(comContext.Trim('=')))
                {
                    if (!(double.TryParse(baseContext.Trim('='), out double baseValue) && double.TryParse(comContext.Trim('='), out double comValue) && Math.Abs(baseValue - comValue) < 1e-12))
                    {
                        SetCellSafe(BaseSheet, baseLocation.Row, col, $"{baseContext} => {comContext}");
                        result = false;
                    }
                }
            }
            return result;
        }
    }
}
