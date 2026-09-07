using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class PinMapSheetComparer : IgxlSheetComparerBase
    {
        public PinMapSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Group Name", ["Comment"])
        {
            KeyColHeaders = ["Group Name", "Pin Name"];
            SheetType = EnumSheetType.DTPinMap;
        }
    }
}
