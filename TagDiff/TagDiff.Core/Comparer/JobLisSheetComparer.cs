using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;

using TagDiff.Core.Common;

namespace TagDiff.Core.Comparer
{
    public class JobLisSheetComparer : IgxlSheetComparerBase
    {
        private readonly bool _isGeneratedJobListSheet;

        public JobLisSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Job Name", ["Comment"])
        {
            KeyColHeaders = ["Job Name"];
            SheetType = EnumSheetType.DTJobListSheet;
            string compName = Path.GetFileNameWithoutExtension(CompFile);
            _isGeneratedJobListSheet = SheetName.EqualsIgnoreCase("GeneratedJobListSheet") || compName.EqualsIgnoreCase("GeneratedJobListSheet");
            SheetType = EnumSheetType.DTJobListSheet;
        }

        protected override bool CompareSheetRow(Location baseLocation, Location compLocation, HashSet<int>? skips = null)
        {
            bool result = true;
            for (int col = baseLocation.Col; col <= BaseSheetEndCol; col++)
            {
                if (skips != null && skips.Contains(col))
                {
                    continue;
                }

                string baseContext = GetCellSafe(BaseSheet, baseLocation.Row, col).Trim('=');
                string comContext = GetCellSafe(CompSheet, compLocation.Row, col).Trim('=');
                bool isDifferent;

                if (_isGeneratedJobListSheet)
                {
                    var baseSet = new HashSet<string>(baseContext.Split([','], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()), StringExtensions.IgnoreCase);
                    IEnumerable<string> comArr = comContext.Split([','], StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim());
                    isDifferent = !baseSet.SetEquals(comArr);
                }
                else
                {
                    isDifferent = !baseContext.EqualsIgnoreCase(comContext);
                }

                if (isDifferent)
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
