using BinCutScriptLib.Algorithm.GradeSearch;
using BinCutScriptLib.Base;
using BinCutScriptLib.Comparer.BVChecker.GetVoltage;

using CommonLib.Extension;

using IgxlLib.Enums;

using TestPlanLib.BinCut.Flow;

namespace BinCutScriptLib.Comparer.BVComparer
{
    public class BvComparerNew : BvComparerBase
    {
        public BvComparerNew(OneGradeSearch oneGradeSearch, EnumJob enumJob, CheckManager checkManager, int tchCnt, EnumSearchType enumSearchType, EnumPrintType enumPrintType)
            : base(oneGradeSearch, enumJob, checkManager, tchCnt, enumSearchType)
        {
            PrintType = enumPrintType;
            BinCutTableType = InstanceBinCut.CurInstanceName.EndsWithIgnoreCase("_HBV>") ||
                              InstanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV_Cs>") ||
                              InstanceBinCut.CurInstanceName.EndsWithIgnoreCase("HBV_Csharp>") ? EnumBinCutTableType.Hv : EnumBinCutTableType.Lv;
            GetVoltageBase = new GetVoltageNew(enumJob, InstanceBinCut);
        }
    }
}
