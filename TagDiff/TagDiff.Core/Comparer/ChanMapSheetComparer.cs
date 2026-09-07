using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class ChanMapSheetComparer : IgxlSheetComparerBase
    {
        public ChanMapSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Pin Name", ["Comment"])
        {
            KeyColHeaders = ["Pin Name"];
            SheetType = EnumSheetType.DTChanMap;
        }
    }
}
