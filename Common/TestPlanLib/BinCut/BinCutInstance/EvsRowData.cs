using System.Collections.Generic;
using System.Linq;

using TestPlanLib.EVS;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class EvsRowData
    {
        public List<EvsCondition> EvsConditions = [];
        public string EvsCategory = "";
        public string EvsRampingCount = "";
        public string EvsStressTime = "";
        public string EvsCoolingTime = "";
        public string EvsPwrPin1 = "";
        public string EvsPwrPin2 = "";
        public string EvsParallelSetting = "";
        public string EvsTotalPwrLimit = "";
        public string EvsType = "";
        public string DvsCoolingPatIndex = "";
        public string DvsCoolingPatSetLoopNumber = "";

        public EvsRowData Copy()
        {
            return new EvsRowData
            {
                EvsConditions = [.. EvsConditions.Select(x => x.Copy())],
                EvsCategory = EvsCategory,
                EvsRampingCount = EvsRampingCount,
                EvsStressTime = EvsStressTime,
                EvsCoolingTime = EvsCoolingTime,
                EvsPwrPin1 = EvsPwrPin1,
                EvsPwrPin2 = EvsPwrPin2,
                EvsParallelSetting = EvsParallelSetting,
                EvsTotalPwrLimit = EvsTotalPwrLimit,
                EvsType = EvsType,
                DvsCoolingPatIndex = DvsCoolingPatIndex,
                DvsCoolingPatSetLoopNumber = DvsCoolingPatSetLoopNumber
            };
        }
    }
}
