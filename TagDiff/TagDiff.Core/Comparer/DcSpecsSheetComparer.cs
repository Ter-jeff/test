using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

using IgxlLib.Enums;
using IgxlLib.IgxlBase;

namespace TagDiff.Core.Comparer
{
    public class DcSpecsSheetComparer : AcSpecsSheetComparer
    {
        public DcSpecsSheetComparer(string sheetName, string baseFile, string compFile, List<InstanceRow> instanceRows) : base(sheetName, baseFile, compFile, instanceRows)
        {
            UseList = new HashSet<string>([.. instanceRows.Select(x => x.DcCategory).Distinct()], StringExtensions.IgnoreCase);
            SheetType = EnumSheetType.DTDCSpecSheet;
        }
    }
}
