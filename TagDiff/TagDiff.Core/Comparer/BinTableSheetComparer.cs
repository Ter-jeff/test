using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlSheets;

namespace TagDiff.Core.Comparer
{
    public class BinTableSheetComparer : IgxlSheetComparerBase
    {
        public BinTableSheetComparer(string sheetName, string baseFile, string compFile, List<SubFlowSheet> subFlowSheets) : base(sheetName, baseFile, compFile, "Name", ["Sort", "Bin", "Comment"])
        {
            UseList = new HashSet<string>([.. subFlowSheets.SelectMany(x => x.Rows).Where(x => x.Opcode.EqualsIgnoreCase("BinTable")).Select(x => x.Parameter)], StringExtensions.IgnoreCase);
            KeyColHeaders = ["Name"];
            SheetType = EnumSheetType.DTBintablesSheet;
        }
    }
}
