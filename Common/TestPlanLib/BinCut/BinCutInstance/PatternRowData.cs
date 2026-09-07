using System.Collections.Generic;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class PatternRowData
    {
        public string PatternPinGroup = "";
        public string PinGroupBinoutFlag = "";
        public string MultiFstpEnable = "";
        public string UserFunction = "";
        public string PerfMode = "";
        public float PmOrder = -1;
        public string CallFlow = "";
        public List<string> PatternList = [];
        public List<string> InitList = [];
        public Dictionary<int, string> PatWithIndex = [];
        public List<string> PayloadList = [];

        public PatternRowData Copy()
        {
            return new PatternRowData
            {
                PatternPinGroup = PatternPinGroup,
                PinGroupBinoutFlag = PinGroupBinoutFlag,
                MultiFstpEnable = MultiFstpEnable,
                UserFunction = UserFunction,
                PerfMode = PerfMode,
                PmOrder = PmOrder,
                CallFlow = CallFlow,
                PatternList = [.. PatternList],
                InitList = [.. InitList],
                PatWithIndex = new Dictionary<int, string>(PatWithIndex),
                PayloadList = [.. PayloadList]
            };
        }
    }
}
