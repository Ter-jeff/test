using System.Collections.Generic;

using Automation.GenerateIgxl.BistBira.Base;

using IgxlLib.IgxlBase;

using TestPlanLib.Basic;

namespace Automation.GenerateIgxl.BistBira.NewLogicData
{
    public class MbistDataStore
    {
        // input variables
        public string Job = "";
        public BistNaming BistNaming = null;
        public Dictionary<string, PatternData> PatternSet = null;

        // Middle variables
        public Dictionary<string, List<PatSet>> DicPatSets = new Dictionary<string, List<PatSet>>();
    }
}
