using System.Collections.Generic;

namespace TestPlanLib.Basic
{
    public class SubprogramMappingSheet
    {
        public List<SubprogramSetting> SubprogramSettings = [];
    }

    public class SubprogramSetting
    {
        public string SubprogramName;
        public List<string> EnableSubFlows = [];
        public SubprogramSetting(string subprogramName, List<string> enableSubFlows)
        {
            SubprogramName = subprogramName;
            EnableSubFlows.AddRange(enableSubFlows);
        }
    }
}
