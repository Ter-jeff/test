using System.Collections.Generic;
using System.Data;
using System.Linq;

using Automation.Const;
using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiAutoZTtr : ContiBase
    {
        public ContiAutoZTtr(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
        {
            ContiFailFlags = new Dictionary<string, string>
            {
                { DcContiConst.FlagNameOpen, DcContiConst.FlagNameOpen },
                { DcContiConst.FlagNameShort, DcContiConst.FlagNameShort },
            };
            if (dcTestContiRow.ConditionDict.TryGetValue("openFlagName", out string flag))
            {
                ContiFailFlags[DcContiConst.FlagNameOpen] = flag;
            }
            if (dcTestContiRow.ConditionDict.TryGetValue("shortFlagName", out flag))
            {
                ContiFailFlags[DcContiConst.FlagNameShort] = flag;
            }
            CombinedBinName = "Bin_DC_" + string.Join("_", ContiFailFlags.Values.Select(x => DcContiConst.FailFlagRegex.Replace(x, "")));
        }

        public override List<FlowRow> GenerateFlowRows()
        {
            return null;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {

            InstanceRow row = new InstanceRow
            {
                TestName = !string.IsNullOrEmpty(DcTestContiRow.InstanceName) ? DcTestContiRow.InstanceName : "DC_Continuity_IONeg_AUTOZ",
                DcCategory = SetCategoryFromCondition(),
                DcSelector = "Typ",
                PinLevels = GetLevels()
            };
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.CSharpFuncNameAutoZ, "conti", true);
            GenerateCSharpInstanceRow(ref function, "");
            row.VbtName = function.FullFunctionName;
            row.VbtType = function.Type;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            return new List<InstanceRow> { row };
        }

        public void GenerateCSharpInstanceRow(ref Function function, string job)
        {
            function.SetParamValue("measuredPins", "CONTINUITY_NEG_AUTOZ");
            function.SetParamValue("forceValue", "-0.0002");
            function.SetParamValue("testLimitMode", "0");
            function.SetParamValue("disconnectPN", "TRUE");
            function.SetParamValue("connectPins", "All_Digital");
            function.SetParamValue("openFlagName", ContiFailFlags[DcContiConst.FlagNameOpen]);
            function.SetParamValue("shortFlagName", ContiFailFlags[DcContiConst.FlagNameShort]);
            function.SetParamValue("lowLimit", "-0.85");
            function.SetParamValue("highLimit", "-0.01");
            function.SetParamValue("isParallelTest", "TRUE");
        }
    }
}
