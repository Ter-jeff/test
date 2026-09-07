using System.Collections.Generic;
using System.Data;

namespace Automation.Reader.ConfigFile.NamingRule.Base
{
    public class ScanConfig : ConfigBase
    {
        public List<ScanConfigSplitRule> SplitRuleList { set; get; } = new List<ScanConfigSplitRule>();
        public DataTable PayloadType { set; get; }
        public string SpecialFlagSetting { set; get; } = string.Empty;
        public Dictionary<string, List<RtosConfig>> NewNamingRulesList { set; get; }
    }

    public class ScanConfigSplitRule
    {
        public string SubName { set; get; }
        public string SubRule { set; get; }
    }
}
