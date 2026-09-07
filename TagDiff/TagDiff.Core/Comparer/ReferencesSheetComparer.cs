using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class ReferencesSheetComparer : IgxlSheetComparerBase
    {
        public ReferencesSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "File Path", ["Comment"])
        {
            KeyColHeaders = ["File Path"];
            SheetType = EnumSheetType.DTReferencesSheet;
        }
    }
}
