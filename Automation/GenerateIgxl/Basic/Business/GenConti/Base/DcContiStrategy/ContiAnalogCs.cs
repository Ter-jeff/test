using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

using Automation.Const;
using Automation.Reader;
using Automation.Static;

using IgxlLib.IgxlBase;

using TestPlanLib.VbtLib;

namespace Automation.GenerateIgxl.Basic.Business.GenConti.Base.DcContiStrategy
{
    public class ContiAnalogCs : ContiAnalog
    {
        public ContiAnalogCs(DcTestContiRow dcTestContiRow) : base(dcTestContiRow)
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

        protected override InstanceRow GenUVI80_Continuity(string job)
        {
            InstanceRow row = new InstanceRow();

            var allLowLimit = DcTestContiRow.Limits.Where(p => p.LimitStage.ToUpper() == job.ToUpper()).Select(p => p.LimitValue).ToList();
            var allHiLimit = DcTestContiRow.Limits.Where(p => p.LimitStage == job.ToUpper()).Select(p => p.HiLimitValue).ToList();
            string differentialPins = DcTestContiRow.Condition.Split(';').FirstOrDefault(x => x.Trim().StartsWith("differentialPins:", StringComparison.OrdinalIgnoreCase));
            row.TestName = DcTestContiRow.InstanceName + "_" + job;
            Function function = TestProgram.VbtFunctionLib.GetFunctionByName(FuncNameConst.CSharpFuncNameAutoZContinuity, "conti", true);
            if (!function.IsFound || function.Type == "VBT")
            {
                return base.GenUVI80_Continuity(job);
            }

            row.DcCategory = "Conti_X_X_X";
            row.DcSelector = "Typ";
            row.PinLevels = GetLevels();

            function.SetParamValue("measuredPins", TestPinGroup);
            function.SetParamValue("forceValue", ForceCondition.ForceValue);
            function.SetParamValue("lowLimit", string.Join(";", allLowLimit));
            function.SetParamValue("highLimit", string.Join(";", allHiLimit));
            function.SetParamValue("testLimitMode", "0");
            function.SetParamValue("disconnectPN", "TRUE");
            function.SetParamValue("openFlagName", ContiFailFlags[DcContiConst.FlagNameOpen]);
            function.SetParamValue("shortFlagName", ContiFailFlags[DcContiConst.FlagNameShort]);
            function.SetParamValue("connectPins", "All_Digital");
            if (!string.IsNullOrEmpty(differentialPins))
            {

                differentialPins = Regex.Replace(differentialPins.Trim(), "^differentialPins:", "", RegexOptions.IgnoreCase);
                function.SetParamValue("differentialPins", differentialPins);
            }
            SetArgumentFromCondition(ref function);

            row.VbtName = function.FullFunctionName;
            row.VbtType = function.Type;
            row.ArgList = function.Parameters;
            row.Args = function.ArgList;
            return row;
        }
    }
}
