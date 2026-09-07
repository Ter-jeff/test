using System.Collections.Generic;

using TestPlanLib.EVS;

namespace TestPlanLib.BinCut.BinCutInstance
{
    internal sealed class EvsColumnIndices
    {
        public List<EvsConditionColIdx> ConditionColIdxDic { get; } = [];
        public int EvsVoltageColNumber { get; set; } = -1;
        public int EvsRampingColNumber { get; set; } = -1;
        public int EvsStressTimeColNumber { get; set; } = -1;
        public int EvsCoolingTimeColNumber { get; set; } = -1;
        public int EvsPwrPin1ColNumber { get; set; } = -1;
        public int EvsPwrPin2ColNumber { get; set; } = -1;
        public int EvsParallelSettingColNumber { get; set; } = -1;
        public int EvsTotalPwrLimitColNumber { get; set; } = -1;
        public int TypeColNumber { get; set; } = -1;
        public int DvsCoolingPatIndexColNumber { get; set; } = -1;
        public int DvsCoolingPatSetLoopNumberColNumber { get; set; } = -1;
    }
}
