using System.Collections.Generic;

namespace TestPlanLib.BinNumber
{
    public class BinNumUsedStatus
    {
        public bool Used { get; set; } = false;
        public List<BinNumInfo> UsedBinNumInfos { get; set; } = [];
    }
}
