using System.Collections.Generic;

namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public class HardIpSheet
    {
        public string SheetName { get; set; }
        public string SubBlock { get; set; }
        public List<PatternRow> PatternRows { get; set; } = new List<PatternRow>();
        public List<HardIpPattern> Rows { get; set; } = new List<HardIpPattern>();

        public string ForceStr { get; set; }
        public int ForceIndex { get; set; }
        public Dictionary<string, int> PlanHeaderIdx { get; set; } = new Dictionary<string, int>();
    }
}
