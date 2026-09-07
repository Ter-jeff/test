using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class PatternSubroutineSheetComparer : IgxlSheetComparerBase
    {
        public PatternSubroutineSheetComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Pattern Filename", ["Comment"])
        {
            KeyColHeaders = ["Pattern Filename"];
            SheetType = EnumSheetType.DTPatternSubroutineSheet;
        }
    }
}
