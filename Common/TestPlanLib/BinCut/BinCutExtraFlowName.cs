using System.Collections.Generic;
using System.Linq;

using CommonLib.Extension;

namespace TestPlanLib.BinCut
{
    public class BinCutExtraFlowName
    {
        public static List<string> GetExtraPerformanceMode(List<string> modeList)
        {
            string last = modeList[0].Split('_').First();
            var extra = new List<string>();
            var flow = new List<string>();
            foreach (string mode in modeList)
            {
                string flowName = mode.Split('_').First();
                if (!flow.Exists(x => x.EqualsIgnoreCase(flowName)))
                {
                    flow.Add(flowName);
                    last = flowName;
                }
                else
                {
                    if (!flowName.EqualsIgnoreCase(last))
                    {
                        extra.Add(mode);
                        flow.Add(mode);
                        last = mode;
                    }
                    else
                    {
                        last = flowName;
                    }
                }
            }
            return extra;
        }
    }
}
