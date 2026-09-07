using System.Collections.Generic;

namespace TestPlanLib.BinCut.BinCutInstance
{
    public class MiscRowData
    {
        public string IsBinout = "";
        public string SplitKey = "";
        public string AcSpec = "";
        public string RetentionWaitTime = "";
        public List<int> RetentionWaitIdx = [];
        public BincutInstanceType Type;
        public string SpecialSetting = "";
        public string Overlay = "";
        public string Char = "";
        public string FunctionName = "";
        public string IsHarvestingCache = "";
        public string EnableCoreMaskCache = "";

        public MiscRowData Copy()
        {
            return new MiscRowData
            {
                IsBinout = IsBinout,
                SplitKey = SplitKey,
                AcSpec = AcSpec,
                RetentionWaitTime = RetentionWaitTime,
                RetentionWaitIdx = [.. RetentionWaitIdx],
                Type = Type,
                SpecialSetting = SpecialSetting,
                Overlay = Overlay,
                Char = Char,
                FunctionName = FunctionName,
                IsHarvestingCache = IsHarvestingCache,
                EnableCoreMaskCache = EnableCoreMaskCache
            };
        }
    }
}
