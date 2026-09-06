using System.Text.RegularExpressions;

using TestPlanLib.BinCut.Flow;

namespace TestPlanLib.BinCut.FlowNew
{
    internal static partial class NewBincutFlowSheetRegex
    {
        [GeneratedRegex(BinCutFlowTable.RegexPerformance, RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
        public static partial Regex Performance();
        [GeneratedRegex("HIP", RegexOptions.IgnoreCase, "en-US")]
        public static partial Regex Hip();
    }
}
