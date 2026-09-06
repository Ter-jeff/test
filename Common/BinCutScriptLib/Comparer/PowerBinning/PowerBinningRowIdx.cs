using System.Collections.Generic;

namespace BinCutScriptLib.Comparer.PowerBinning
{
    public class PowerBinningRowIdx
    {
        public int Site;
        public int RowIdx;
        public bool BPass;
        public List<Dictionary<string, bool>> CheckSheetList = [];
    }
}
