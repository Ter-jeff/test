using System.Collections.Generic;

namespace Automation.Reader.ConfigFile.NamingRule.Base
{
    public class MbistConfig : ConfigBase
    {
        public Dictionary<string, RtosConfig> NamingRulesList { set; get; } = new Dictionary<string, RtosConfig>();
        public string SpecialFlagSetting { set; get; } = string.Empty;
        public Dictionary<string, List<KeyAndPosition>> RepairFlagSetting { set; get; } = new Dictionary<string, List<KeyAndPosition>>();
        public Dictionary<string, MbistProductionSpecialNamingRule> SpecialNamingRules { set; get; } = new Dictionary<string, MbistProductionSpecialNamingRule>();
    }

    public class KeyAndPosition
    {
        public string KeyPositions { set; get; }
        public string Keys { set; get; }
    }

    public class MbistProductionSpecialNamingRule
    {
        public string UseLabelPosition { set; get; }
        public bool VMarginNeedNumber { set; get; }
        public string FlagUsePatternPosition { get; set; } = string.Empty;
    }
}
