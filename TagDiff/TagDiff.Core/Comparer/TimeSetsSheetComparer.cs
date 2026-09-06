using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class TimeSetsSheetComparer : IgxlSheetComparerBase
    {
        public TimeSetsSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Time Set", [])
        {
            KeyColHeaders = ["Time Set", "Name"];
            SheetType = EnumSheetType.DTTimesetBasicSheet;
        }
    }
}
