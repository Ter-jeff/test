using System.Collections.Generic;
using System.IO;
using System.Linq;

using IgxlLib.Enums;

using TagDiff.Core.Common;
using TagDiff.Core.Output;

using TagDiffCore.Utility;

namespace TagDiff.Core.Comparer
{
    public class OtherSheetComparer(string sheetName, string baseFile, string compFile) : SheetComparerBase(sheetName, baseFile, compFile, "", [])
    {
        public override List<CompareReport> Compare()
        {
            var result = new CompareReport();
            if (CompSheet.Count == 0)
            {
                result.SheetName = SheetName;
                result.SheetType = EnumSheetType.DTUnknown;
                result.TotalCount = BaseSheet.Count;
                result.ManualCount = BaseSheet.Count;
                result.SemiAutoCount = 0;
                return [result];
            }

            int totalCount = 0;
            int failCount = 0;
            int semiAutoCount = 0;
            for (int row = BaseSheetStartRow; row <= BaseSheetEndRow; row++)
            {
                if (SheetProvider.IsEmptyRow(BaseSheet, row))
                {
                    continue;
                }

                totalCount++;
                if (!CompareSheetRow(new Location(row, 1), new Location(row, 1)))
                {
                    semiAutoCount++;
                }
            }

            IEnumerable<string> lines = BaseSheet.Select(row => string.Join("\t", row));
            File.WriteAllLines(BaseFile, lines);

            result.SheetName = SheetName;
            result.SheetType = EnumSheetType.DTUnknown;
            result.TotalCount = totalCount;
            result.ManualCount = failCount;
            result.SemiAutoCount = semiAutoCount;

            return [result];
        }
    }
}
