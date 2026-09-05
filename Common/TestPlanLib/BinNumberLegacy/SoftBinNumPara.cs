using System.Collections.Generic;

namespace TestPlanLib.BinNumberLegacy
{
    public class SoftBinNumPara
    {
        public string Category { set; get; } = "";
        public string Module { set; get; } = "";
        public string SubModule { set; get; } = "";
        public string Block { set; get; } = "";
        public string Level { set; get; } = "";
        public string PerformanceMode { set; get; } = "";
        public Dictionary<string, string> ColumnContentDic { set; get; } = [];
    }
}
