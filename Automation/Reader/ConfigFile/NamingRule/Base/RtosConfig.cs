using System.Collections.Generic;

namespace Automation.Reader.ConfigFile.NamingRule.Base
{
    public class RtosConfig : ConfigBase
    {
        public List<int> NamingRuleInitSequenceList = new List<int>();
        public List<string> NamingRuleInitList = new List<string>();
        public List<string> NamingRulePayloadList = new List<string>();
        public string Block { get; set; }
        public string Module { get; set; }
    }
}
