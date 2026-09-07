using IgxlLib.Enums;

namespace TagDiff.Core.Comparer
{
    public class BinCutEqnComparer : IgxlSheetComparerBase
    {
        public BinCutEqnComparer(string sheetName, string baseFile, string compFile) : base(sheetName, baseFile, compFile, "Binned", ["SoftBin", "HardBin"])
        {
            KeyColHeaders = ["Mode_EQN"];
            SheetType = EnumSheetType.BinCut_eqn;
        }
    }
}
