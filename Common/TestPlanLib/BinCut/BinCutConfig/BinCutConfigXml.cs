using System;
using System.Collections.Generic;

namespace TestPlanLib.BinCut.BinCutConfig
{
    public class BinCutConfigXml
    {
        public Dictionary<string, EnumPowerType> PowerTypeDic { get; set; } = [];
        public List<Tuple<string, string>> IdsNameList { get; set; } = [];
        public Dictionary<string, string> DomainInOtherRail2PowerDic { get; set; } = [];
        public Dictionary<string, int> Dssc { get; set; } = [];
        public Dictionary<string, bool> PowerBinningConfig { get; set; } = [];
        public string SafeVoltageFollowPayload { get; set; } = string.Empty;
    }
}
