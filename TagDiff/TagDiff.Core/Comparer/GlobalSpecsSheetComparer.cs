using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class GlobalSpecsSheetComparer : IgxlSheetComparerBase
    {
        public GlobalSpecsSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Symbol", ["Comment"])
        {
            KeyColHeaders = ["Symbol", "Job"];
            SheetType = EnumSheetType.DTGlobalSpecSheet;
        }
    }
}
