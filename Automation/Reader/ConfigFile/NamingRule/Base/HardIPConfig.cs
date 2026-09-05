using System.Collections.Generic;
using System.Data;

namespace Automation.Reader.ConfigFile.NamingRule.Base
{
    public class HardIpConfig : ConfigBase
    {
        public List<HardIpConfigSplitRule> SplitRuleList { set; get; } = new List<HardIpConfigSplitRule>();
        public Dictionary<string, RtosConfig> NamingRulesList { set; get; } = new Dictionary<string, RtosConfig>();
        public DataTable PayloadType { set; get; }
        public string SpecialFlagSetting { set; get; } = string.Empty;
    }

    public class HardIpConfigSplitRule
    {
        public string SubName { set; get; }
        public string SubRule { set; get; }
    }
}
