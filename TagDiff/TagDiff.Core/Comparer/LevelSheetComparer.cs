using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class LevelSheetComparer : IgxlSheetComparerBase
    {
        public LevelSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Pin/Group", [])
        {
            KeyColHeaders = ["Pin/Group", "Parameter"];
            SheetType = EnumSheetType.DTLevelSheet;
        }
    }
}
